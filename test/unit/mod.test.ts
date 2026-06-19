import { assertEquals, assertStringIncludes } from 'https://deno.land/std@0.208.0/assert/mod.ts';
import { tools } from '../../mod.ts';
import type { PluginContext } from 'cortex/plugins';

const mockContext: PluginContext = {
  pluginId: 'cortex-plugin-teams',
  pluginDir: '/tmp/plugins/cortex-plugin-teams',
  state: {
    get: async () => null,
    set: async () => {},
  },
  config: {},
};

function findTool(name: string) {
  return tools.find((t) => t.definition.name === name);
}

Deno.test('teams_send_message - rejects missing required params', async () => {
  const tool = findTool('teams_send_message');
  if (!tool) throw new Error('teams_send_message tool not found');

  const result = await tool.execute({}, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'required');
});

Deno.test('teams_send_message - rejects invalid target_type', async () => {
  const tool = findTool('teams_send_message');
  if (!tool) throw new Error('teams_send_message tool not found');

  const result = await tool.execute({
    target_id: 'ch123',
    target_type: 'invalid',
    content: 'Hello',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'must be');
});

Deno.test('teams_send_message - rejects invalid adaptive_card JSON', async () => {
  const tool = findTool('teams_send_message');
  if (!tool) throw new Error('teams_send_message tool not found');

  const result = await tool.execute({
    target_id: 'ch123',
    target_type: 'channel',
    content: 'Hello',
    adaptive_card: 'not-json',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'valid JSON');
});

Deno.test('teams_send_message - rejects missing API config', async () => {
  const tool = findTool('teams_send_message');
  if (!tool) throw new Error('teams_send_message tool not found');

  const result = await tool.execute({
    target_id: 'ch123',
    target_type: 'channel',
    content: 'Hello',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('teams_read_messages - rejects missing required params', async () => {
  const tool = findTool('teams_read_messages');
  if (!tool) throw new Error('teams_read_messages tool not found');

  const result = await tool.execute({}, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'required');
});

Deno.test('teams_read_messages - rejects invalid target_type', async () => {
  const tool = findTool('teams_read_messages');
  if (!tool) throw new Error('teams_read_messages tool not found');

  const result = await tool.execute({
    target_id: 'ch123',
    target_type: 'invalid',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'must be');
});

Deno.test('teams_read_messages - rejects missing API config', async () => {
  const tool = findTool('teams_read_messages');
  if (!tool) throw new Error('teams_read_messages tool not found');

  const result = await tool.execute({
    target_id: 'ch123',
    target_type: 'channel',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('teams_list_channels - rejects missing team_id', async () => {
  const tool = findTool('teams_list_channels');
  if (!tool) throw new Error('teams_list_channels tool not found');

  const result = await tool.execute({}, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'team_id');
});

Deno.test('teams_list_channels - rejects missing API config', async () => {
  const tool = findTool('teams_list_channels');
  if (!tool) throw new Error('teams_list_channels tool not found');

  const result = await tool.execute({ team_id: 'team123' }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('teams_create_meeting - rejects missing required params', async () => {
  const tool = findTool('teams_create_meeting');
  if (!tool) throw new Error('teams_create_meeting tool not found');

  const result = await tool.execute({}, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'required');
});

Deno.test('teams_create_meeting - rejects missing API config', async () => {
  const tool = findTool('teams_create_meeting');
  if (!tool) throw new Error('teams_create_meeting tool not found');

  const result = await tool.execute({
    subject: 'Sprint Review',
    start_time: '2025-01-01T09:00:00',
    end_time: '2025-01-01T10:00:00',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('teams_create_meeting - accepts optional attendees', async () => {
  const tool = findTool('teams_create_meeting');
  if (!tool) throw new Error('teams_create_meeting tool not found');

  const result = await tool.execute({
    subject: 'Sprint Review',
    start_time: '2025-01-01T09:00:00',
    end_time: '2025-01-01T10:00:00',
    attendees: 'alice@example.com,bob@example.com',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('teams_send_notification - rejects missing required params', async () => {
  const tool = findTool('teams_send_notification');
  if (!tool) throw new Error('teams_send_notification tool not found');

  const result = await tool.execute({}, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'required');
});

Deno.test('teams_send_notification - rejects invalid urgency', async () => {
  const tool = findTool('teams_send_notification');
  if (!tool) throw new Error('teams_send_notification tool not found');

  const result = await tool.execute({
    user_id: 'user123',
    message: 'Hello',
    urgency: 'invalid',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'must be one of');
});

Deno.test('teams_send_notification - rejects missing API config', async () => {
  const tool = findTool('teams_send_notification');
  if (!tool) throw new Error('teams_send_notification tool not found');

  const result = await tool.execute({
    user_id: 'user123',
    message: 'Hello',
    urgency: 'normal',
  }, mockContext);
  assertEquals(result.success, false);
  assertStringIncludes(result.error, 'not configured');
});

Deno.test('tools array exported', () => {
  assertEquals(tools.length, 5);
  assertEquals(tools[0].definition.name, 'teams_send_message');
  assertEquals(tools[1].definition.name, 'teams_read_messages');
  assertEquals(tools[2].definition.name, 'teams_list_channels');
  assertEquals(tools[3].definition.name, 'teams_create_meeting');
  assertEquals(tools[4].definition.name, 'teams_send_notification');
});
