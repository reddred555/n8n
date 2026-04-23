import { Logger } from '@n8n/backend-common';
import { mockInstance } from '@n8n/backend-test-utils';
import type express from 'express';
import { mock, type MockProxy } from 'jest-mock-extended';
import * as n8nCore from 'n8n-core';
import { BinaryDataService, ErrorReporter } from 'n8n-core';
import type {
	Workflow,
	INode,
	IDataObject,
	IWebhookResponseData,
	IDeferredPromise,
	IN8nHttpFullResponse,
	IWorkflowBase,
	IRunExecutionData,
	IExecuteData,
	IWebhookData,
} from 'n8n-workflow';
import {
	createDeferredPromise,
	FORM_NODE_TYPE,
	WAIT_NODE_TYPE,
	CHAT_TRIGGER_NODE_TYPE,
	WorkflowConfigurationError,
	NodeOperationError,
	MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
} from 'n8n-workflow';
import type { Readable } from 'stream';
import { finished } from 'stream/promises';

import {
	autoDetectResponseMode,
	handleFormRedirectionCase,
	setupResponseNodePromise,
	prepareExecutionData,
	handleHostedChatResponse,
	executeWebhook,
	_privateGetWebhookErrorMessage,
} from '../webhook-helpers';
import type { IWebhookResponseCallbackData, WebhookRequest } from '../webhook.types';

import { ActiveExecutions } from '@/active-executions';
import { AuthService } from '@/auth/auth.service';
import { EventService } from '@/events/event.service';
import { OwnershipService } from '@/services/ownership.service';
import { WorkflowStatisticsService } from '@/services/workflow-statistics.service';
import * as WorkflowExecuteAdditionalData from '@/workflow-execute-additional-data';
import { WorkflowRunner } from '@/workflow-runner';
import { WebhookService } from '../webhook.service';

jest.mock('stream/promises', () => ({
	finished: jest.fn(),
}));

describe('autoDetectResponseMode', () => {
	let workflow: MockProxy<Workflow>;

	beforeEach(() => {
		workflow = mock<Workflow>();
		workflow.nodes = {};
	});

	test('should return hostedChat when start node is CHAT_TRIGGER_NODE_TYPE, method is POST, and public is true', () => {
		const workflowStartNode = mock<INode>({
			type: CHAT_TRIGGER_NODE_TYPE,
			parameters: { options: { responseMode: 'responseNodes' } },
		});
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'POST');
		expect(result).toBe('hostedChat');
	});

	test('should return undefined if start node is WAIT_NODE_TYPE with resume not equal to form', () => {
		const workflowStartNode = mock<INode>({
			type: WAIT_NODE_TYPE,
			parameters: { resume: 'webhook' },
		});
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'POST');
		expect(result).toBeUndefined();
	});

	test('should return responseNode when start node is FORM_NODE_TYPE and method is POST', () => {
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			name: 'startNode',
			parameters: {},
		});
		workflow.getChildNodes.mockReturnValue(['childNode']);
		workflow.nodes.childNode = mock<INode>({
			type: WAIT_NODE_TYPE,
			parameters: { resume: 'form' },
			disabled: false,
		});
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'POST');
		expect(result).toBe('responseNode');
	});

	test('should return formPage when start node is FORM_NODE_TYPE and method is POST and there is a following FORM_NODE_TYPE node', () => {
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			name: 'startNode',
			parameters: {},
		});
		workflow.getChildNodes.mockReturnValue(['childNode']);
		workflow.nodes.childNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {
				operation: 'completion',
			},
			disabled: false,
		});
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'POST');
		expect(result).toBe('formPage');
	});

	test('should return undefined when start node is FORM_NODE_TYPE with no other form child nodes', () => {
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			name: 'startNode',
			parameters: {},
		});
		workflow.getChildNodes.mockReturnValue([]);
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'POST');
		expect(result).toBeUndefined();
	});

	test('should return undefined for non-matching node type and method', () => {
		const workflowStartNode = mock<INode>({ type: 'someOtherNodeType', parameters: {} });
		const result = autoDetectResponseMode(workflowStartNode, workflow, 'GET');
		expect(result).toBeUndefined();
	});
});

describe('handleFormRedirectionCase', () => {
	test('should return data unchanged if start node is WAIT_NODE_TYPE with resume not equal to form', () => {
		const data: IWebhookResponseCallbackData = {
			responseCode: 302,
			headers: { location: 'http://example.com' },
		};
		const workflowStartNode = mock<INode>({
			type: WAIT_NODE_TYPE,
			parameters: { resume: 'webhook' },
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result).toEqual(data);
	});

	test('should modify data if start node type matches and responseCode is a redirect', () => {
		const data: IWebhookResponseCallbackData = {
			responseCode: 302,
			headers: { location: 'http://example.com' },
		};
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {},
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result.responseCode).toBe(200);
		expect(result.data).toEqual({ redirectURL: 'http://example.com' });
		expect((result?.headers as IDataObject)?.location).toBeUndefined();
	});

	test('should not modify data if location header is missing', () => {
		const data: IWebhookResponseCallbackData = { responseCode: 302, headers: {} };
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {},
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result).toEqual(data);
	});

	test('should block javascript: URLs for security', () => {
		const data: IWebhookResponseCallbackData = {
			responseCode: 302,
			headers: { location: 'javascript:alert(document.domain)' },
		};
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {},
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result.responseCode).toBe(200);
		expect(result.data).toBeUndefined();
		expect((result?.headers as IDataObject)?.location).toBeUndefined();
	});

	test('should block data: URLs for security', () => {
		const data: IWebhookResponseCallbackData = {
			responseCode: 302,
			headers: { location: 'data:text/html,<script>alert(1)</script>' },
		};
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {},
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result.responseCode).toBe(200);
		expect(result.data).toBeUndefined();
		expect((result?.headers as IDataObject)?.location).toBeUndefined();
	});

	test('should allow https: URLs', () => {
		const data: IWebhookResponseCallbackData = {
			responseCode: 302,
			headers: { location: 'https://example.com/callback' },
		};
		const workflowStartNode = mock<INode>({
			type: FORM_NODE_TYPE,
			parameters: {},
		});
		const result = handleFormRedirectionCase(data, workflowStartNode);
		expect(result.responseCode).toBe(200);
		expect(result.data).toEqual({ redirectURL: 'https://example.com/callback' });
	});
});

describe('setupResponseNodePromise', () => {
	const workflowId = 'test-workflow-id';
	const executionId = 'test-execution-id';
	const res = mock<express.Response>();
	const responseCallback = jest.fn();
	const workflowStartNode = mock<INode>();
	const workflow = mock<Workflow>({ id: workflowId });
	const binaryDataService = mockInstance(BinaryDataService);
	const errorReporter = mockInstance(ErrorReporter);
	const logger = mockInstance(Logger);

	let responsePromise: IDeferredPromise<IN8nHttpFullResponse>;

	beforeEach(() => {
		jest.resetAllMocks();

		responsePromise = createDeferredPromise<IN8nHttpFullResponse>();

		res.header.mockReturnValue(res);
		res.end.mockReturnValue(res);
	});

	test('should handle regular response object', async () => {
		setupResponseNodePromise(
			responsePromise,
			res,
			responseCallback,
			workflowStartNode,
			executionId,
			workflow,
		);

		responsePromise.resolve({
			body: { data: 'test data' },
			headers: { 'content-type': 'application/json' },
			statusCode: 200,
		});
		await new Promise(process.nextTick);

		expect(responseCallback).toHaveBeenCalledWith(null, {
			data: { data: 'test data' },
			headers: { 'content-type': 'application/json' },
			responseCode: 200,
		});
		expect(res.end).toHaveBeenCalled();
	});

	test('should handle binary data with ID', async () => {
		const mockStream = mock<Readable>();
		binaryDataService.getAsStream.mockResolvedValue(mockStream);

		setupResponseNodePromise(
			responsePromise,
			res,
			responseCallback,
			workflowStartNode,
			executionId,
			workflow,
		);

		responsePromise.resolve({
			body: { binaryData: { id: 'binary-123' } },
			headers: { 'content-type': 'image/jpeg' },
			statusCode: 200,
		});
		await new Promise(process.nextTick);

		expect(binaryDataService.getAsStream).toHaveBeenCalledWith('binary-123');
		expect(res.setHeaders).toHaveBeenCalledWith(new Map([['content-type', 'image/jpeg']]));
		expect(mockStream.pipe).toHaveBeenCalledWith(res, { end: false });
		expect(finished).toHaveBeenCalledWith(mockStream);
		expect(responseCallback).toHaveBeenCalledWith(null, { noWebhookResponse: true });
	});

	test('should handle buffer response', async () => {
		setupResponseNodePromise(
			responsePromise,
			res,
			responseCallback,
			workflowStartNode,
			executionId,
			workflow,
		);

		const buffer = Buffer.from('test buffer');
		responsePromise.resolve({
			body: buffer,
			headers: { 'content-type': 'text/plain' },
			statusCode: 200,
		});
		await new Promise(process.nextTick);

		expect(res.setHeaders).toHaveBeenCalledWith(new Map([['content-type', 'text/plain']]));
		expect(res.end).toHaveBeenCalledWith(buffer);
		expect(responseCallback).toHaveBeenCalledWith(null, { noWebhookResponse: true });
	});

	test('should handle errors properly', async () => {
		setupResponseNodePromise(
			responsePromise,
			res,
			responseCallback,
			workflowStartNode,
			executionId,
			workflow,
		);

		const error = new Error('Test error');
		responsePromise.reject(error);
		await new Promise(process.nextTick);

		expect(errorReporter.error).toHaveBeenCalledWith(error);
		expect(logger.error).toHaveBeenCalledWith(
			`Error with Webhook-Response for execution "${executionId}": "${error.message}"`,
			{ executionId, workflowId },
		);
		expect(responseCallback).toHaveBeenCalledWith(error, {});
	});
});

describe('handleHostedChatResponse', () => {
	it('should send executionStarted: true, executionId, and resumeToken when responseMode is hostedChat', async () => {
		const res = {
			send: jest.fn(),
			end: jest.fn(),
		} as unknown as express.Response;
		const responseMode = 'hostedChat';
		let didSendResponse = false;
		const executionId = '123';
		const resumeToken = 'a'.repeat(64);

		const result = handleHostedChatResponse(
			res,
			responseMode,
			didSendResponse,
			executionId,
			resumeToken,
		);

		expect(res.send).toHaveBeenCalledWith({ executionStarted: true, executionId, resumeToken });
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(res.end).toHaveBeenCalled();
		expect(result).toBe(true);
	});

	it('should not send response when responseMode is not hostedChat', () => {
		const res = {
			send: jest.fn(),
			end: jest.fn(),
		} as unknown as express.Response;
		const executionId = 'testExecutionId';
		let didSendResponse = false;
		const responseMode = 'responseNode';

		const result = handleHostedChatResponse(res, responseMode, didSendResponse, executionId);

		expect(res.send).not.toHaveBeenCalled();
		expect(res.end).not.toHaveBeenCalled();
		expect(result).toBe(false);
	});

	it('should not send response when didSendResponse is true', () => {
		const res = {
			send: jest.fn(),
			end: jest.fn(),
		} as unknown as express.Response;
		const executionId = 'testExecutionId';
		let didSendResponse = true;
		const responseMode = 'hostedChat';

		const result = handleHostedChatResponse(res, responseMode, didSendResponse, executionId);

		expect(res.send).not.toHaveBeenCalled();
		expect(res.end).not.toHaveBeenCalled();
		expect(result).toBe(true);
	});
});

describe('prepareExecutionData', () => {
	const workflowStartNode = mock<INode>({ name: 'Start' });
	const webhookResultData: IWebhookResponseData = {
		workflowData: [[{ json: { data: 'test' } }]],
	};
	const workflowData = mock<IWorkflowBase>({
		id: 'workflow1',
		pinData: { nodeA: [{ json: { pinned: true } }] },
	});

	test('should create new execution data when not provided', () => {
		const { runExecutionData, pinData } = prepareExecutionData(
			'manual',
			workflowStartNode,
			webhookResultData,
			undefined,
		);

		const nodeExecuteData = runExecutionData.executionData?.nodeExecutionStack?.[0];
		expect(nodeExecuteData).toBeDefined();
		expect(nodeExecuteData?.node).toBe(workflowStartNode);
		expect(nodeExecuteData?.data.main).toBe(webhookResultData.workflowData);
		expect(pinData).toBeUndefined();
	});

	test('should update existing runExecutionData when executionId is defined', () => {
		const executionId = 'test-execution-id';
		const nodeExecutionStack: IExecuteData[] = [
			{
				node: workflowStartNode,
				data: { main: [[{ json: { oldData: true } }]] },
				source: null,
			},
		];
		const existingRunExecutionData = {
			startData: {},
			resultData: { runData: {} },
			executionData: {
				contextData: {},
				nodeExecutionStack,
				waitingExecution: {},
			},
		} as IRunExecutionData;

		prepareExecutionData(
			'manual',
			workflowStartNode,
			webhookResultData,
			existingRunExecutionData,
			undefined,
			undefined,
			executionId,
		);

		expect(nodeExecutionStack[0]?.data.main).toBe(webhookResultData.workflowData);
	});

	test('should set destination node when provided', () => {
		const { runExecutionData } = prepareExecutionData(
			'manual',
			workflowStartNode,
			webhookResultData,
			undefined,
			{},
			{ nodeName: 'targetNode', mode: 'inclusive' },
		);

		expect(runExecutionData.startData?.destinationNode).toEqual({
			nodeName: 'targetNode',
			mode: 'inclusive',
		});
	});

	test('should update execution data with execution data merge', () => {
		const runExecutionDataMerge = {
			resultData: {
				error: { message: 'Test error' },
			},
		};

		const { runExecutionData } = prepareExecutionData(
			'manual',
			workflowStartNode,
			webhookResultData,
			undefined,
			runExecutionDataMerge,
		);

		expect(runExecutionData.resultData.error).toEqual({ message: 'Test error' });
	});

	test('should set pinData when execution mode is manual', () => {
		const { runExecutionData, pinData } = prepareExecutionData(
			'manual',
			workflowStartNode,
			webhookResultData,
			undefined,
			{},
			undefined,
			undefined,
			workflowData,
		);

		expect(pinData).toBe(workflowData.pinData);
		expect(runExecutionData.resultData.pinData).toBe(workflowData.pinData);
	});

	test('should not set pinData when execution mode is not manual or evaluation', () => {
		const { runExecutionData, pinData } = prepareExecutionData(
			'webhook',
			workflowStartNode,
			webhookResultData,
			undefined,
			{},
			undefined,
			undefined,
			workflowData,
		);

		expect(pinData).toBeUndefined();
		expect(runExecutionData.resultData.pinData).toBeUndefined();
	});

	describe('MICROSOFT_AGENT365_TRIGGER_NODE_TYPE merge condition', () => {
		test('should merge nodeExecutionStack when node type is MICROSOFT_AGENT365_TRIGGER_NODE_TYPE and runExecutionData exists', () => {
			const microsoftAgentNode = mock<INode>({
				name: 'Microsoft Agent 365',
				type: MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			});

			const existingNodeExecutionStack: IExecuteData[] = [
				{
					node: mock<INode>({ name: 'ExistingNode' }),
					data: {
						main: [[{ json: { existing: 'data' } }]],
					},
					source: null,
				},
			];

			const existingRunExecutionData: IRunExecutionData = {
				version: 1,
				startData: {},
				resultData: { runData: {} },
				executionData: {
					contextData: {},
					metadata: {},
					nodeExecutionStack: existingNodeExecutionStack,
					waitingExecution: {},
					waitingExecutionSource: {},
				},
			} as IRunExecutionData;

			const { runExecutionData } = prepareExecutionData(
				'trigger',
				microsoftAgentNode,
				webhookResultData,
				existingRunExecutionData,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack).toHaveLength(1);
			expect(runExecutionData.executionData?.nodeExecutionStack[0].node.name).toBe(
				'Microsoft Agent 365',
			);
			expect(runExecutionData.executionData?.nodeExecutionStack[0].node.type).toBe(
				MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			);
			expect(runExecutionData.executionData?.nodeExecutionStack[0].data.main[0]).toHaveLength(1);
			expect(runExecutionData.executionData?.nodeExecutionStack[0].data.main[0]?.[0]?.json).toEqual(
				{
					existing: 'data',
					data: 'test',
				},
			);
		});

		test('should not merge when node type is MICROSOFT_AGENT365_TRIGGER_NODE_TYPE but runExecutionData is undefined', () => {
			const microsoftAgentNode = mock<INode>({
				name: 'Microsoft Agent 365',
				type: MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			});

			const { runExecutionData } = prepareExecutionData(
				'trigger',
				microsoftAgentNode,
				webhookResultData,
				undefined,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack).toHaveLength(1);
			expect(runExecutionData.executionData?.nodeExecutionStack[0].node).toEqual(
				microsoftAgentNode,
			);
		});

		test('should not merge when node type is MICROSOFT_AGENT365_TRIGGER_NODE_TYPE but nodeExecutionStack is undefined', () => {
			const microsoftAgentNode = mock<INode>({
				name: 'Microsoft Agent 365',
				type: MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			});

			const existingRunExecutionData: IRunExecutionData = {
				version: 1,
				startData: {},
				resultData: { runData: {} },
				executionData: {
					contextData: {},
					metadata: {},
					nodeExecutionStack: undefined as any,
					waitingExecution: {},
					waitingExecutionSource: {},
				},
			} as IRunExecutionData;

			const { runExecutionData } = prepareExecutionData(
				'trigger',
				microsoftAgentNode,
				webhookResultData,
				existingRunExecutionData,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack).toBeUndefined();
		});

		test('should not merge when node type is not MICROSOFT_AGENT365_TRIGGER_NODE_TYPE', () => {
			const regularNode = mock<INode>({
				name: 'Regular Webhook',
				type: 'n8n-nodes-base.webhook',
			});

			const existingNodeExecutionStack: IExecuteData[] = [
				{
					node: mock<INode>({ name: 'ExistingNode' }),
					data: {
						main: [[{ json: { existing: 'data' } }]],
					},
					source: null,
				},
			];

			const existingRunExecutionData: IRunExecutionData = {
				version: 1,
				startData: {},
				resultData: { runData: {} },
				executionData: {
					contextData: {},
					metadata: {},
					nodeExecutionStack: existingNodeExecutionStack,
					waitingExecution: {},
					waitingExecutionSource: {},
				},
			} as IRunExecutionData;

			const { runExecutionData } = prepareExecutionData(
				'trigger',
				regularNode,
				webhookResultData,
				existingRunExecutionData,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack).toHaveLength(1);
			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.node.name).toBe(
				'ExistingNode',
			);

			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.data.main).toEqual([
				[{ json: { existing: 'data' } }],
			]);
		});

		test('should merge existing data with new data for MICROSOFT_AGENT365_TRIGGER_NODE_TYPE', () => {
			const microsoftAgentNode = mock<INode>({
				name: 'Microsoft Agent 365',
				type: MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			});

			const existingData: IExecuteData = {
				node: mock<INode>({ name: 'ExistingNode' }),
				data: {
					main: [[{ json: { existing: 'preserved' } }]],
				},
				source: { main: [{ previousNode: 'test' }] },
			};

			const existingRunExecutionData: IRunExecutionData = {
				version: 1,
				startData: {},
				resultData: { runData: {} },
				executionData: {
					contextData: {},
					metadata: {},
					nodeExecutionStack: [existingData],
					waitingExecution: {},
					waitingExecutionSource: {},
				},
			} as IRunExecutionData;

			const { runExecutionData } = prepareExecutionData(
				'trigger',
				microsoftAgentNode,
				webhookResultData,
				existingRunExecutionData,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack).toHaveLength(1);

			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.node.name).toBe(
				'Microsoft Agent 365',
			);
			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.node.type).toBe(
				MICROSOFT_AGENT365_TRIGGER_NODE_TYPE,
			);

			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.data.main[0]).toHaveLength(1);
			expect(
				runExecutionData.executionData?.nodeExecutionStack?.[0]?.data.main[0]?.[0]?.json,
			).toEqual({
				existing: 'preserved',
				data: 'test',
			});

			expect(runExecutionData.executionData?.nodeExecutionStack?.[0]?.source).toBeNull();
		});
	});
});

describe('getWebhookErrorMessage', () => {
	const workflowStartNode = mock<INode>({ name: 'Start' });
	it('should surface WorkflowConfigurationError', () => {
		const err = new WorkflowConfigurationError(workflowStartNode, new Error('test'));
		expect(_privateGetWebhookErrorMessage(err, 'Webhook')).toEqual(err.message);
	});

	it('should obfuscate other errors', () => {
		const err = new NodeOperationError(workflowStartNode, new Error('test'));
		expect(_privateGetWebhookErrorMessage(err, 'Webhook')).toContain(
			'Error: Workflow could not be started',
		);
	});
});

describe('executeWebhook - context establishment ordering', () => {
	const callOrder: string[] = [];
	let workflowRunner: ReturnType<typeof mockInstance<WorkflowRunner>>;
	let webhookService: ReturnType<typeof mockInstance<WebhookService>>;
	let activeExecutions: ReturnType<typeof mockInstance<ActiveExecutions>>;
	let establishSpy: jest.SpyInstance;
	let getBaseSpy: jest.SpyInstance;

	beforeAll(() => {
		mockInstance(Logger);
		mockInstance(BinaryDataService);
		mockInstance(ErrorReporter);
		mockInstance(AuthService);
		mockInstance(EventService);
		mockInstance(WorkflowStatisticsService);
	});

	beforeEach(() => {
		jest.restoreAllMocks();
		callOrder.length = 0;

		workflowRunner = mockInstance(WorkflowRunner);
		webhookService = mockInstance(WebhookService);
		activeExecutions = mockInstance(ActiveExecutions);
		const ownershipService = mockInstance(OwnershipService);

		ownershipService.getWorkflowProjectCached.mockResolvedValue({
			id: 'project-1',
		} as never);

		webhookService.runWebhook.mockImplementation(async () => ({
			workflowData: [[{ json: { headers: { authorization: 'Bearer token' } } }]],
		}));

		workflowRunner.run.mockImplementation(async () => {
			callOrder.push('WorkflowRunner.run');
			return 'exec-123';
		});

		activeExecutions.getPostExecutePromise.mockReturnValue(
			new Promise(() => {
				/* never resolves */
			}),
		);

		// Spy on the n8n-core module export that webhook-helpers imports
		establishSpy = jest
			.spyOn(n8nCore, 'establishExecutionContext')
			.mockImplementation(async (_workflow, runExecutionData) => {
				// Simulate real behaviour: set runtimeData so downstream code can assert it
				runExecutionData.executionData!.runtimeData = {
					version: 1,
					establishedAt: Date.now(),
					source: 'webhook',
					redaction: { version: 1, policy: 'none' },
				};
				callOrder.push('establishExecutionContext');
			});

		// Provide a valid additionalData
		getBaseSpy = jest.spyOn(WorkflowExecuteAdditionalData, 'getBase').mockResolvedValue({
			formWaitingBaseUrl: 'http://localhost/form',
			webhookWaitingBaseUrl: 'http://localhost/waiting',
		} as never);
	});

	afterEach(() => {
		establishSpy.mockRestore();
		getBaseSpy.mockRestore();
	});

	const buildFixtures = () => {
		const workflowStartNode = mock<INode>({
			name: 'Webhook',
			type: 'n8n-nodes-base.webhook',
			typeVersion: 2,
			parameters: {},
		});

		const webhookData = mock<IWebhookData>({
			node: 'Webhook',
			workflowId: 'wf-1',
			webhookDescription: {
				responseMode: '={{$parameter["responseMode"]}}',
				responseCode: '={{$parameter["responseCode"]}}',
				responseData: '={{$parameter["responseData"]}}',
				responsePropertyName: undefined,
				responseContentType: undefined,
				responseBinaryPropertyName: undefined,
				responseHeaders: undefined,
			} as never,
		});

		const expression = {
			getSimpleParameterValue: jest.fn((_node, _expr, _mode, _keys, _runIndex, fallback) => {
				// Return 'onReceived' for responseMode, 200 for responseCode, fallback otherwise
				if (fallback === 'onReceived') return 'onReceived';
				if (fallback === 200) return 200;
				return fallback;
			}),
			getComplexParameterValue: jest.fn(
				(_node, _expr, _mode, _keys, _runIndex, fallback) => fallback,
			),
		};

		const workflow = mock<Workflow>({
			id: 'wf-1',
			name: 'Test Workflow',
		});
		// Attach properties not covered by the mock proxy
		(workflow as unknown as { expression: typeof expression }).expression = expression;
		(workflow as unknown as { nodeTypes: unknown }).nodeTypes = {
			getByNameAndVersion: jest.fn().mockReturnValue({
				description: { name: 'webhook', properties: [] },
			}),
		};
		workflow.getChildNodes.mockReturnValue([]);

		const workflowData = mock<IWorkflowBase>({
			id: 'wf-1',
			name: 'Test Workflow',
			nodes: [],
			connections: {},
		});

		const req = mock<WebhookRequest>({
			method: 'POST',
			headers: { authorization: 'Bearer token' },
			params: {},
		});
		(req as unknown as { contentType: string }).contentType = 'application/json';
		(req as unknown as { body: unknown }).body = {};
		(req as unknown as { query: unknown }).query = {};

		const res = mock<express.Response>();
		(res as unknown as { headersSent: boolean }).headersSent = false;

		const responseCallback = jest.fn();

		return { workflow, workflowStartNode, webhookData, workflowData, req, res, responseCallback };
	};

	it('calls establishExecutionContext before WorkflowRunner.run', async () => {
		const { workflow, workflowStartNode, webhookData, workflowData, req, res, responseCallback } =
			buildFixtures();

		await executeWebhook(
			workflow,
			webhookData,
			workflowData,
			workflowStartNode,
			'webhook',
			undefined,
			undefined,
			undefined,
			req,
			res,
			responseCallback,
		);

		expect(establishSpy).toHaveBeenCalledTimes(1);
		expect(workflowRunner.run).toHaveBeenCalledTimes(1);
		expect(callOrder).toEqual(['establishExecutionContext', 'WorkflowRunner.run']);
	});

	it('passes runExecutionData with runtimeData populated to WorkflowRunner.run', async () => {
		const { workflow, workflowStartNode, webhookData, workflowData, req, res, responseCallback } =
			buildFixtures();

		let capturedRunData: IRunExecutionData | undefined;
		workflowRunner.run.mockImplementation(async (data) => {
			capturedRunData = data.executionData;
			callOrder.push('WorkflowRunner.run');
			return 'exec-456';
		});

		await executeWebhook(
			workflow,
			webhookData,
			workflowData,
			workflowStartNode,
			'webhook',
			undefined,
			undefined,
			undefined,
			req,
			res,
			responseCallback,
		);

		expect(capturedRunData?.executionData?.runtimeData).toBeDefined();
		expect(capturedRunData?.executionData?.runtimeData?.source).toBe('webhook');
	});

	it('invokes establishExecutionContext with the prepared runExecutionData', async () => {
		const { workflow, workflowStartNode, webhookData, workflowData, req, res, responseCallback } =
			buildFixtures();

		await executeWebhook(
			workflow,
			webhookData,
			workflowData,
			workflowStartNode,
			'webhook',
			undefined,
			undefined,
			undefined,
			req,
			res,
			responseCallback,
		);

		const [, runExecutionDataArg] = establishSpy.mock.calls[0];
		expect(runExecutionDataArg.executionData?.nodeExecutionStack?.[0]?.node).toBe(
			workflowStartNode,
		);
		// Trigger items from webhookResultData.workflowData are attached to the stack
		expect(runExecutionDataArg.executionData?.nodeExecutionStack?.[0]?.data?.main?.[0]).toEqual([
			{ json: { headers: { authorization: 'Bearer token' } } },
		]);
	});
});
