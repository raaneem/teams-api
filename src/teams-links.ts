/**
 * Teams deep links.
 *
 * See docs/findings.md ("Message search"): message permalinks have the form
 * `https://teams.microsoft.com/l/message/{conversationId}/{messageId}`, with
 * `parentMessageId` for replies in a channel thread.
 */

export function buildTeamsMessageLink(
  conversationId: string,
  messageId: string,
  parentMessageId?: string | null,
): string {
  const parent = parentMessageId
    ? `?parentMessageId=${encodeURIComponent(parentMessageId)}`
    : "";
  return `https://teams.microsoft.com/l/message/${encodeURIComponent(conversationId)}/${encodeURIComponent(messageId)}${parent}`;
}
