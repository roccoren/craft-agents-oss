import type { ModelRegistry as PiModelRegistry } from '@earendil-works/pi-coding-agent';

/**
 * GitHub Copilot models that the account's live `/models` API serves (and that
 * Craft's dynamic model picker surfaces via `drivers/pi.ts:fetchCopilotModels`)
 * but that the pinned Pi SDK static catalog (`github-copilot.models`) does NOT
 * yet include. Without a runtime registry entry, `resolvePiModel` returns
 * `undefined` for these ids and selecting them fails, even though the picker
 * shows them.
 *
 * This bridges the gap until the Pi SDK catalog catches up: when the auth
 * provider is `github-copilot`, we merge these into the registry's
 * github-copilot provider (copying request shape from an existing sibling model
 * where available). Add new ids here whenever GitHub Copilot ships a model
 * ahead of the SDK. Entries already present in the SDK catalog are skipped, so
 * this list is safe to keep even after the SDK adds them.
 */
export interface CopilotExtraModel {
  id: string;
  name: string;
  contextWindow: number;
  maxTokens: number;
}

export const GITHUB_COPILOT_EXTRA_MODELS: CopilotExtraModel[] = [
  { id: 'gpt-5.6-sol', name: 'GPT-5.6 Sol', contextWindow: 1_050_000, maxTokens: 128_000 },
  { id: 'gpt-5.6-terra', name: 'GPT-5.6 Terra', contextWindow: 1_050_000, maxTokens: 128_000 },
  { id: 'gpt-5.6-luna', name: 'GPT-5.6 Luna', contextWindow: 1_050_000, maxTokens: 128_000 },
];

/** Headers that identify us as a VS Code Copilot client (same as the Pi SDK catalog). */
const COPILOT_HEADERS: Record<string, string> = {
  'User-Agent': 'GitHubCopilotChat/0.35.0',
  'Editor-Version': 'vscode/1.107.0',
  'Editor-Plugin-Version': 'copilot-chat/0.35.0',
  'Copilot-Integration-Id': 'vscode-chat',
};

/** Extract the per-account API base URL from a Copilot token's `proxy-ep` field. */
function baseUrlFromCopilotToken(token: string | undefined): string | undefined {
  if (!token) return undefined;
  const match = token.match(/proxy-ep=([^;]+)/);
  if (!match?.[1]) return undefined;
  return `https://${match[1].replace(/^proxy\./, 'api.')}`;
}

type LooseModel = {
  id: string;
  name: string;
  api?: string;
  baseUrl?: string;
  reasoning?: boolean;
  thinkingLevelMap?: unknown;
  input?: ('text' | 'image')[];
  cost?: unknown;
  contextWindow?: number;
  maxTokens?: number;
  headers?: Record<string, string>;
  compat?: unknown;
  provider?: string;
};

/**
 * Register any {@link GITHUB_COPILOT_EXTRA_MODELS} that are missing from the
 * registry into the `github-copilot` provider, preserving the existing models.
 *
 * `registerProvider` fully replaces a provider's models, so the existing
 * github-copilot models are re-supplied alongside the extras. Per-request auth
 * still resolves from `authStorage` first (kept fresh by `token_update`), so the
 * `apiKey` passed here is only a validation stub / fallback.
 */
export function augmentGithubCopilotRegistry(
  registry: PiModelRegistry,
  copilotAccessToken: string | undefined,
): void {
  const existing = (registry.getAll() as unknown as LooseModel[]).filter(
    (m) => m.provider === 'github-copilot',
  );
  const presentIds = new Set(existing.map((m) => m.id));
  const missing = GITHUB_COPILOT_EXTRA_MODELS.filter((m) => !presentIds.has(m.id));
  if (missing.length === 0) return;

  // registerProvider requires an apiKey (or oauth) when defining models.
  if (!copilotAccessToken) return;

  const sibling =
    existing.find((m) => m.id === 'gpt-5.5') ??
    existing.find((m) => m.api === 'openai-responses') ??
    existing[0];

  const baseUrl =
    sibling?.baseUrl ??
    baseUrlFromCopilotToken(copilotAccessToken) ??
    'https://api.individual.githubcopilot.com';

  const api = (sibling?.api as string | undefined) ?? 'openai-responses';
  const headers = sibling?.headers ?? { ...COPILOT_HEADERS };
  const thinkingLevelMap = sibling?.thinkingLevelMap;
  const cost = sibling?.cost ?? { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 };
  const compat = sibling?.compat;

  const toDef = (m: LooseModel) => ({
    id: m.id,
    name: m.name,
    api: m.api,
    baseUrl: m.baseUrl,
    reasoning: m.reasoning ?? false,
    thinkingLevelMap: m.thinkingLevelMap,
    input: m.input ?? ['text'],
    cost: m.cost,
    contextWindow: m.contextWindow ?? 128_000,
    maxTokens: m.maxTokens ?? 16_384,
    headers: m.headers,
    compat: m.compat,
  });

  const extraDefs = missing.map((m) => ({
    id: m.id,
    name: m.name,
    api,
    baseUrl,
    reasoning: true,
    thinkingLevelMap,
    input: ['text', 'image'] as ('text' | 'image')[],
    cost,
    contextWindow: m.contextWindow,
    maxTokens: m.maxTokens,
    headers,
    compat,
  }));

  registry.registerProvider('github-copilot', {
    baseUrl,
    apiKey: copilotAccessToken,
    models: [...existing.map(toDef), ...extraDefs] as never,
  });
}
