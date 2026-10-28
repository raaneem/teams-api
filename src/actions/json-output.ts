/**
 * Stable JSON shapes for `--format json`.
 *
 * `formatOutput` wraps these in a versioned envelope:
 *   { "schemaVersion": 1, "ok": true, "command": "...", "data": ..., "meta": {...} }
 *
 * Unlike `--format detailed` (the raw internal objects), these shapes are a
 * contract for scripts and bots: fields are only added, never renamed, and
 * missing values are `null` rather than absent.
 */

import type {
  Conversation,
  Member,
  Message,
  MessageSearchResult,
} from "../types.js";
import { extractQuote } from "./formatters.js";
import { extractMriSuffix } from "../mri.js";
import { buildTeamsMessageLink } from "../teams-links.js";

// ── Result context ───────────────────────────────────────────────────
//
// Actions resolve details (such as the conversation ID behind `--chat`) that
// the JSON output needs but the result value does not carry. They attach the
// details here instead of changing the result, so concise and detailed output
// stay exactly as before.

export interface ResultContext {
  conversationId?: string;
  conversationLabel?: string;
  threadRootId?: string | null;
  since?: string | null;
  order?: string;
  limitReached?: boolean;
}

const resultContexts = new WeakMap<object, ResultContext>();

export function attachResultContext<Result>(
  result: Result,
  context: ResultContext,
): Result {
  if (typeof result === "object" && result !== null) {
    resultContexts.set(result, context);
  }
  return result;
}

export function getResultContext(result: unknown): ResultContext {
  return typeof result === "object" && result !== null
    ? (resultContexts.get(result) ?? {})
    : {};
}

// ── Shapes ───────────────────────────────────────────────────────────

export function conversationToJson(conversation: Conversation) {
  return {
    id: conversation.id,
    topic: conversation.topic || null,
    type: conversation.threadType,
    lastMessageAt: conversation.lastMessageTime,
    memberCount: conversation.memberCount,
  };
}

export function messageToJson(
  message: Message,
  context: Pick<ResultContext, "conversationId" | "threadRootId">,
) {
  const conversationId = context.conversationId ?? null;
  const threadRootId = context.threadRootId ?? null;
  const { quote, body } = extractQuote(message.content);
  return {
    id: message.id,
    conversationId,
    threadRootId,
    link:
      conversationId && message.id
        ? buildTeamsMessageLink(conversationId, message.id, threadRootId)
        : null,
    createdAt: message.originalArrivalTime || null,
    editedAt: message.editTime,
    sender: {
      id: extractMriSuffix(message.senderMri) || null,
      name: message.senderDisplayName || null,
    },
    type: message.messageType,
    subject: message.subject,
    text: body,
    quote:
      quote || message.quotedMessageId
        ? { messageId: message.quotedMessageId, text: quote }
        : null,
    html: message.content,
    isDeleted: message.isDeleted,
    mentions: message.mentions.map((mention) => ({
      id: mention.id,
      name: mention.displayName,
    })),
    reactions: message.reactions.map((reaction) => ({
      key: reaction.key,
      count: reaction.users.length,
      users: reaction.users.map((user) => ({
        id: user.mri,
        name: user.displayName || null,
      })),
    })),
    files: message.files.map((file) => ({
      name: file.fileName,
      type: file.fileType,
      url: file.shareUrl || file.fileUrl || null,
    })),
    images: message.images.map((image) => ({
      amsObjectId: image.amsObjectId,
      width: image.width,
      height: image.height,
    })),
  };
}

/** Latest `createdAt` among the messages, as ISO-8601, or null. */
export function newestCreatedAt(
  messages: ReadonlyArray<{ createdAt: string | null }>,
): string | null {
  let newestTime = Number.NEGATIVE_INFINITY;
  let newest: string | null = null;
  for (const message of messages) {
    const time = message.createdAt ? Date.parse(message.createdAt) : NaN;
    if (!Number.isNaN(time) && time > newestTime) {
      newestTime = time;
      newest = new Date(time).toISOString();
    }
  }
  return newest;
}

export function memberToJson(member: Member) {
  return {
    id: member.id,
    name: member.displayName || null,
    role: member.role,
    type: member.memberType,
  };
}

export function searchResultToJson(result: MessageSearchResult) {
  return {
    id: result.messageId ?? null,
    searchResultId: result.id ?? null,
    conversationId: result.conversationId ?? null,
    conversationName: result.conversationName ?? null,
    threadRootId: result.parentMessageId ?? null,
    threadType: result.threadType ?? null,
    link: result.webUrl ?? null,
    createdAt: result.sentAt ?? null,
    sender: {
      id: result.senderMri ?? null,
      name: result.senderDisplayName ?? null,
      email: result.senderEmail ?? null,
    },
    subject: result.subject ?? null,
    preview: result.preview ?? null,
  };
}
