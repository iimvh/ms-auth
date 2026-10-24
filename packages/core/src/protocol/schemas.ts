import { z } from "zod/v4";

export const SignalRHandshakeResponse = z.object({
  error: z.string().optional(),
});

export const ThrottlingInfo = z.object({
  maxNumUserMessagesInConversation: z.number(),
  numUserMessagesInConversation: z.number(),
  numLongDocSummaryUserMessagesInConversation: z.number(),
});

export const ClassifierScore = z.object({
  component: z.string(),
  score: z.number(),
});

export const ContentGenerationProgress = z.object({
  contentType: z.string().optional(),
  size: z.string().optional(),
  orientation: z.string().optional(),
  ImageReferenceUrls: z.array(z.string()).optional(),
  fileToken: z.string().optional(),
  pollUrl: z.string().optional(),
  status: z.number().optional(),
});

export const BotMessage = z.object({
  text: z.string(),
  author: z.literal("bot"),
  contentType: z.string().optional(),
  contentGenerationProgressList: z.array(ContentGenerationProgress).optional(),
  responseIdentifier: z.string().optional(),
  createdAt: z.string().optional(),
  timestamp: z.string().optional(),
  messageId: z.string().optional(),
  requestId: z.string().optional(),
  offense: z.string().optional(),
  messageType: z.string().optional(),
  adaptiveCards: z.array(z.any()).optional(),
  sourceAttributions: z.array(z.any()).optional(),
  contentOrigin: z.string().optional(),
  scores: z.array(ClassifierScore).optional(),
  turnCount: z.number().optional(),
  turnState: z.string().optional(),
  suggestedResponses: z
    .array(
      z.object({
        text: z.string(),
        commandText: z.string().optional(),
        author: z.string().optional(),
      }),
    )
    .optional(),
});

export const CursorInfo = z.object({
  j: z.string(),
  p: z.number(),
});

export const DeltaUpdate = z.object({
  writeAtCursor: z.string(),
  sourceAttributions: z.array(z.any()).optional(),
  streamingMode: z.literal("Delta").optional(),
  nonce: z.string().optional(),
});

export const MessageUpdate = z.object({
  messages: z.array(BotMessage.partial().extend({ author: z.string() })),
  nonce: z.string().optional(),
  cursor: CursorInfo.optional(),
  streamingMode: z.string().optional(),
  requestId: z.string().optional(),
});

export const ThrottlingUpdate = z.object({
  nonce: z.string().optional(),
  requestId: z.string().optional(),
  throttling: ThrottlingInfo,
});

export const CompletionFrame = z.object({
  type: z.literal(3),
  invocationId: z.string(),
  error: z.string().optional(),
  result: z.any().optional(),
});

export const CloseFrame = z.object({
  type: z.literal(7),
  error: z.string().optional(),
});

export const JwtClaims = z.object({
  aud: z.string(),
  iss: z.string(),
  oid: z.string(),
  tid: z.string(),
  exp: z.number(),
  name: z.string().optional(),
  upn: z.string().optional(),
});
