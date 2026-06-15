# Microsoft Teams Integration

Cortex agent as a Microsoft Teams bot.

## Installation

```bash
cortex plugin install github:CortexPrism/cortex-plugin-teams
```

## Tools

### teams_send_message
Send a message to a channel or chat.
- `target_id` (string, required) — Channel or chat ID
- `target_type` (enum: channel, chat, required) — Target type
- `content` (string, required) — Message content
- `adaptive_card` (string, optional) — JSON adaptive card

### teams_read_messages
Read messages from a channel or chat.
- `target_id` (string, required) — Channel or chat ID
- `target_type` (enum: channel, chat, required) — Target type
- `limit` (number, default: 20) — Maximum messages

### teams_list_channels
List team channels.
- `team_id` (string, required) — Team ID

### teams_create_meeting
Create a Teams meeting.
- `subject` (string, required) — Meeting subject
- `start_time` (string, required) — Start time (ISO 8601)
- `end_time` (string, required) — End time (ISO 8601)
- `attendees` (string, optional) — Comma-separated emails

### teams_send_notification
Send a notification to a user.
- `user_id` (string, required) — Target user ID
- `message` (string, required) — Notification message
- `urgency` (enum: normal, high, urgent, required) — Urgency level

## Configuration

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| teamsTenantId | text | Yes | Microsoft Teams tenant ID |
| teamsClientId | text | Yes | Entra ID application client ID |
| teamsClientSecret | secret | Yes | Entra ID application client secret |

## License

MIT
