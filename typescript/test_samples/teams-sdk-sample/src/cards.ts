// Adaptive cards used by the sample. Constructed as plain JSON objects to
// avoid pulling in @microsoft/teams.cards — the wire format is the same.

export function helpCard() {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'Teams SDK Feature Showcase', weight: 'Bolder', size: 'Large' },
      {
        type: 'TextBlock',
        text: 'Teams SDK handlers (TEAMS_APP)',
        weight: 'Bolder',
        spacing: 'Medium',
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'help', value: 'This command list' },
          { title: 'react', value: 'Bot adds/removes emoji reactions' },
          { title: 'quote', value: 'Bot quotes its own message' },
          { title: 'targeted', value: 'Ephemeral message visible only to sender' },
          { title: 'task', value: 'Task module fetch/submit flow' },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Agents SDK fallthrough handlers (AGENT_SDK_APP)',
        weight: 'Bolder',
        spacing: 'Medium',
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'agents sdk react', value: "Reach teams.ts's API client from an Agents SDK handler" },
          { title: 'agents sdk proactive', value: 'Trigger a proactive send from an Agents SDK handler' },
          { title: 'channel', value: 'Report the channel this turn arrived on and how it was routed' },
          { title: 'anything else', value: "Echo via Agents SDK '[Agent SDK] You said: ...'" },
        ],
      },
    ],
  } as const;
}

export function taskLauncherCard() {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: '📋 Task module demo', weight: 'Bolder' },
      {
        type: 'TextBlock',
        text: 'Press the button to open a task module.',
        wrap: true,
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: 'Open task module',
        data: { msteams: { type: 'task/fetch' }, dialog_id: 'open_task' },
      },
    ],
  } as const;
}

export function taskFormCard() {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'Tell us something:', weight: 'Bolder' },
      { type: 'Input.Text', id: 'note', placeholder: 'Type here…' },
    ],
    actions: [{ type: 'Action.Submit', title: 'Submit' }],
  } as const;
}
