import type { PluginContext, Tool, ToolCallResult, ToolContext } from './types.ts';

let pluginConfig: Record<string, unknown> = {};

export async function onLoad(ctx: PluginContext): Promise<void> {
  ctx.logger.info(`[cortex-plugin-teams] Loaded`);
  pluginConfig = await ctx.config.get() as Record<string, unknown>;
}

export async function onUnload(_ctx: PluginContext): Promise<void> {}

const teamsSendMessageTool: Tool = {
  definition: {
    name: 'teams_send_message',
    description: 'Send a message to a channel or chat',
    params: [
      {
        name: 'target_id',
        type: 'string',
        description: 'Target channel or chat ID',
        required: true,
      },
      {
        name: 'target_type',
        type: 'string',
        description: 'Type of target',
        required: true,
        enum: ['channel', 'chat'],
      },
      { name: 'content', type: 'string', description: 'Message content', required: true },
      {
        name: 'adaptive_card',
        type: 'string',
        description: 'JSON string of adaptive card',
        required: false,
      },
    ],
    capabilities: ['network:fetch'],
  },
  execute: async (args: Record<string, unknown>, _ctx: ToolContext): Promise<ToolCallResult> => {
    const start = Date.now();
    try {
      const targetId = args.target_id as string;
      const targetType = args.target_type as string;
      const content = args.content as string;

      if (!targetId || !targetType || !content) {
        return {
          toolName: 'teams_send_message',
          success: false,
          output: '',
          error: 'target_id, target_type, and content are required',
          durationMs: Date.now() - start,
        };
      }

      if (!['channel', 'chat'].includes(targetType)) {
        return {
          toolName: 'teams_send_message',
          success: false,
          output: '',
          error: 'target_type must be "channel" or "chat"',
          durationMs: Date.now() - start,
        };
      }

      const accessToken = pluginConfig.teamsClientSecret as string;
      if (!accessToken) {
        return {
          toolName: 'teams_send_message',
          success: false,
          output: '',
          error: 'Teams not configured. Set teamsTenantId, teamsClientId, and teamsClientSecret.',
          durationMs: Date.now() - start,
        };
      }

      const messageBody: Record<string, unknown> = {
        body: { content, contentType: 'text' },
      };

      if (args.adaptive_card) {
        try {
          messageBody.attachments = [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: JSON.parse(args.adaptive_card as string),
          }];
        } catch {
          return {
            toolName: 'teams_send_message',
            success: false,
            output: '',
            error: 'adaptive_card must be valid JSON',
            durationMs: Date.now() - start,
          };
        }
      }

      let endpoint: string;
      if (targetType === 'channel') {
        endpoint = `https://graph.microsoft.com/v1.0/teams/${
          encodeURIComponent(targetId)
        }/channels/messages`;
      } else {
        endpoint = `https://graph.microsoft.com/v1.0/chats/${
          encodeURIComponent(targetId)
        }/messages`;
      }

      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(messageBody),
      });

      if (!response.ok) {
        return {
          toolName: 'teams_send_message',
          success: false,
          output: '',
          error: `Teams API error: ${response.status}`,
          durationMs: Date.now() - start,
        };
      }

      const data = await response.json();
      return {
        toolName: 'teams_send_message',
        success: true,
        output: JSON.stringify(data),
        durationMs: Date.now() - start,
      };
    } catch (error) {
      return {
        toolName: 'teams_send_message',
        success: false,
        output: '',
        error: `Failed to send message: ${error instanceof Error ? error.message : String(error)}`,
        durationMs: Date.now() - start,
      };
    }
  },
};

const teamsReadMessagesTool: Tool = {
  definition: {
    name: 'teams_read_messages',
    description: 'Read messages from a channel or chat',
    params: [
      {
        name: 'target_id',
        type: 'string',
        description: 'Target channel or chat ID',
        required: true,
      },
      {
        name: 'target_type',
        type: 'string',
        description: 'Type of target',
        required: true,
        enum: ['channel', 'chat'],
      },
      {
        name: 'limit',
        type: 'number',
        description: 'Maximum number of messages',
        required: false,
        default: 20,
      },
    ],
    capabilities: ['network:fetch'],
  },
  execute: async (args: Record<string, unknown>, _ctx: ToolContext): Promise<ToolCallResult> => {
    const start = Date.now();
    try {
      const targetId = args.target_id as string;
      const targetType = args.target_type as string;

      if (!targetId || !targetType) {
        return {
          toolName: 'teams_read_messages',
          success: false,
          output: '',
          error: 'target_id and target_type are required',
          durationMs: Date.now() - start,
        };
      }

      if (!['channel', 'chat'].includes(targetType)) {
        return {
          toolName: 'teams_read_messages',
          success: false,
          output: '',
          error: 'target_type must be "channel" or "chat"',
          durationMs: Date.now() - start,
        };
      }

      const accessToken = pluginConfig.teamsClientSecret as string;
      if (!accessToken) {
        return {
          toolName: 'teams_read_messages',
          success: false,
          output: '',
          error: 'Teams not configured',
          durationMs: Date.now() - start,
        };
      }

      const limit = (args.limit as number) ?? 20;

      let endpoint: string;
      if (targetType === 'channel') {
        endpoint = `https://graph.microsoft.com/v1.0/teams/${
          encodeURIComponent(targetId)
        }/channels/messages?$top=${limit}`;
      } else {
        endpoint = `https://graph.microsoft.com/v1.0/chats/${
          encodeURIComponent(targetId)
        }/messages?$top=${limit}`;
      }

      const response = await fetch(endpoint, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      if (!response.ok) {
        return {
          toolName: 'teams_read_messages',
          success: false,
          output: '',
          error: `Teams API error: ${response.status}`,
          durationMs: Date.now() - start,
        };
      }

      const data = await response.json();
      return {
        toolName: 'teams_read_messages',
        success: true,
        output: JSON.stringify(data),
        durationMs: Date.now() - start,
      };
    } catch (error) {
      return {
        toolName: 'teams_read_messages',
        success: false,
        output: '',
        error: `Failed to read messages: ${error instanceof Error ? error.message : String(error)}`,
        durationMs: Date.now() - start,
      };
    }
  },
};

const teamsListChannelsTool: Tool = {
  definition: {
    name: 'teams_list_channels',
    description: 'List team channels',
    params: [
      { name: 'team_id', type: 'string', description: 'Team ID', required: true },
    ],
    capabilities: ['network:fetch'],
  },
  execute: async (args: Record<string, unknown>, _ctx: ToolContext): Promise<ToolCallResult> => {
    const start = Date.now();
    try {
      const teamId = args.team_id as string;
      if (!teamId) {
        return {
          toolName: 'teams_list_channels',
          success: false,
          output: '',
          error: 'team_id is required',
          durationMs: Date.now() - start,
        };
      }

      const accessToken = pluginConfig.teamsClientSecret as string;
      if (!accessToken) {
        return {
          toolName: 'teams_list_channels',
          success: false,
          output: '',
          error: 'Teams not configured',
          durationMs: Date.now() - start,
        };
      }

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels`,
        { headers: { Authorization: `Bearer ${accessToken}` } },
      );

      if (!response.ok) {
        return {
          toolName: 'teams_list_channels',
          success: false,
          output: '',
          error: `Teams API error: ${response.status}`,
          durationMs: Date.now() - start,
        };
      }

      const data = await response.json();
      return {
        toolName: 'teams_list_channels',
        success: true,
        output: JSON.stringify(data),
        durationMs: Date.now() - start,
      };
    } catch (error) {
      return {
        toolName: 'teams_list_channels',
        success: false,
        output: '',
        error: `Failed to list channels: ${error instanceof Error ? error.message : String(error)}`,
        durationMs: Date.now() - start,
      };
    }
  },
};

const teamsCreateMeetingTool: Tool = {
  definition: {
    name: 'teams_create_meeting',
    description: 'Create a Teams meeting',
    params: [
      { name: 'subject', type: 'string', description: 'Meeting subject', required: true },
      {
        name: 'start_time',
        type: 'string',
        description: 'Start time in ISO 8601 format',
        required: true,
      },
      {
        name: 'end_time',
        type: 'string',
        description: 'End time in ISO 8601 format',
        required: true,
      },
      {
        name: 'attendees',
        type: 'string',
        description: 'Comma-separated email addresses',
        required: false,
      },
    ],
    capabilities: ['network:fetch'],
  },
  execute: async (args: Record<string, unknown>, _ctx: ToolContext): Promise<ToolCallResult> => {
    const start = Date.now();
    try {
      const subject = args.subject as string;
      const startTime = args.start_time as string;
      const endTime = args.end_time as string;

      if (!subject || !startTime || !endTime) {
        return {
          toolName: 'teams_create_meeting',
          success: false,
          output: '',
          error: 'subject, start_time, and end_time are required',
          durationMs: Date.now() - start,
        };
      }

      const accessToken = pluginConfig.teamsClientSecret as string;
      if (!accessToken) {
        return {
          toolName: 'teams_create_meeting',
          success: false,
          output: '',
          error: 'Teams not configured',
          durationMs: Date.now() - start,
        };
      }

      const meetingBody: Record<string, unknown> = {
        subject,
        start: { dateTime: startTime, timeZone: 'UTC' },
        end: { dateTime: endTime, timeZone: 'UTC' },
      };

      if (args.attendees) {
        meetingBody.attendees = (args.attendees as string).split(',').map((email) => ({
          emailAddress: { address: email.trim(), name: email.trim() },
          type: 'required',
        }));
      }

      const response = await fetch(
        'https://graph.microsoft.com/v1.0/me/onlineMeetings',
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
          body: JSON.stringify(meetingBody),
        },
      );

      if (!response.ok) {
        return {
          toolName: 'teams_create_meeting',
          success: false,
          output: '',
          error: `Teams API error: ${response.status}`,
          durationMs: Date.now() - start,
        };
      }

      const data = await response.json();
      return {
        toolName: 'teams_create_meeting',
        success: true,
        output: JSON.stringify(data),
        durationMs: Date.now() - start,
      };
    } catch (error) {
      return {
        toolName: 'teams_create_meeting',
        success: false,
        output: '',
        error: `Failed to create meeting: ${
          error instanceof Error ? error.message : String(error)
        }`,
        durationMs: Date.now() - start,
      };
    }
  },
};

const teamsSendNotificationTool: Tool = {
  definition: {
    name: 'teams_send_notification',
    description: 'Send a notification to a user',
    params: [
      { name: 'user_id', type: 'string', description: 'Target user ID', required: true },
      { name: 'message', type: 'string', description: 'Notification message', required: true },
      {
        name: 'urgency',
        type: 'string',
        description: 'Urgency level',
        required: true,
        enum: ['normal', 'high', 'urgent'],
      },
    ],
    capabilities: ['network:fetch'],
  },
  execute: async (args: Record<string, unknown>, _ctx: ToolContext): Promise<ToolCallResult> => {
    const start = Date.now();
    try {
      const userId = args.user_id as string;
      const message = args.message as string;
      const urgency = args.urgency as string;

      if (!userId || !message || !urgency) {
        return {
          toolName: 'teams_send_notification',
          success: false,
          output: '',
          error: 'user_id, message, and urgency are required',
          durationMs: Date.now() - start,
        };
      }

      if (!['normal', 'high', 'urgent'].includes(urgency)) {
        return {
          toolName: 'teams_send_notification',
          success: false,
          output: '',
          error: 'urgency must be one of: normal, high, urgent',
          durationMs: Date.now() - start,
        };
      }

      const accessToken = pluginConfig.teamsClientSecret as string;
      if (!accessToken) {
        return {
          toolName: 'teams_send_notification',
          success: false,
          output: '',
          error: 'Teams not configured',
          durationMs: Date.now() - start,
        };
      }

      const response = await fetch(
        'https://graph.microsoft.com/v1.0/teams/sendActivityNotification',
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            topic: {
              source: 'text',
              value: `Notification (${urgency})`,
              webUrl: 'https://teams.microsoft.com',
            },
            activityType: 'systemDefault',
            previewText: { content: message },
            recipients: [{ '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId }],
          }),
        },
      );

      if (!response.ok) {
        return {
          toolName: 'teams_send_notification',
          success: false,
          output: '',
          error: `Teams API error: ${response.status}`,
          durationMs: Date.now() - start,
        };
      }

      return {
        toolName: 'teams_send_notification',
        success: true,
        output: `Notification sent to user ${userId}`,
        durationMs: Date.now() - start,
      };
    } catch (error) {
      return {
        toolName: 'teams_send_notification',
        success: false,
        output: '',
        error: `Failed to send notification: ${
          error instanceof Error ? error.message : String(error)
        }`,
        durationMs: Date.now() - start,
      };
    }
  },
};

export const tools: Tool[] = [
  teamsSendMessageTool,
  teamsReadMessagesTool,
  teamsListChannelsTool,
  teamsCreateMeetingTool,
  teamsSendNotificationTool,
];
