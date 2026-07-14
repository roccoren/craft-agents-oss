import { describe, expect, it } from 'bun:test';
import { augmentGithubCopilotRegistry, GITHUB_COPILOT_EXTRA_MODELS } from './github-copilot-extra-models.ts';

type FakeModel = {
  id: string;
  name: string;
  provider: string;
  api?: string;
  baseUrl?: string;
  reasoning?: boolean;
  input?: ('text' | 'image')[];
  cost?: unknown;
  contextWindow?: number;
  maxTokens?: number;
  headers?: Record<string, string>;
};

/** Minimal stand-in for the Pi SDK ModelRegistry covering the methods used. */
function makeFakeRegistry(initial: FakeModel[]) {
  let models = [...initial];
  const calls: Array<{ provider: string; config: { baseUrl?: string; apiKey?: string; models?: FakeModel[] } }> = [];
  const registry = {
    getAll: () => models,
    find: (provider: string, id: string) => models.find(m => m.provider === provider && m.id === id),
    registerProvider: (provider: string, config: { baseUrl?: string; apiKey?: string; models?: Array<Record<string, unknown>> }) => {
      calls.push({ provider, config: config as never });
      if (config.models && config.models.length > 0) {
        models = models.filter(m => m.provider !== provider);
        for (const def of config.models) {
          models.push({ ...(def as FakeModel), provider });
        }
      }
    },
  };
  return { registry, calls, current: () => models };
}

const COPILOT_TOKEN = 'tid=abc;exp=123;proxy-ep=proxy.enterprise.githubcopilot.com;foo=bar';

function siblingModel(): FakeModel {
  return {
    id: 'gpt-5.5',
    name: 'GPT-5.5',
    provider: 'github-copilot',
    api: 'openai-responses',
    baseUrl: 'https://api.enterprise.githubcopilot.com',
    reasoning: true,
    input: ['text', 'image'],
    cost: { input: 5, output: 30, cacheRead: 0.5, cacheWrite: 0 },
    contextWindow: 1_000_000,
    maxTokens: 128_000,
    headers: { 'Copilot-Integration-Id': 'vscode-chat' },
  };
}

describe('augmentGithubCopilotRegistry', () => {
  it('adds the missing gpt-5.6 models while preserving existing ones', () => {
    const { registry } = makeFakeRegistry([siblingModel()]);

    augmentGithubCopilotRegistry(registry as never, COPILOT_TOKEN);

    for (const extra of GITHUB_COPILOT_EXTRA_MODELS) {
      const found = registry.find('github-copilot', extra.id);
      expect(found).toBeDefined();
      expect(found!.api).toBe('openai-responses');
      expect(found!.baseUrl).toBe('https://api.enterprise.githubcopilot.com');
    }
    // Existing sibling still resolvable.
    expect(registry.find('github-copilot', 'gpt-5.5')).toBeDefined();
  });

  it('is a no-op when all extra models already exist in the catalog', () => {
    const withAll: FakeModel[] = [
      siblingModel(),
      ...GITHUB_COPILOT_EXTRA_MODELS.map(m => ({ ...siblingModel(), id: m.id, name: m.name })),
    ];
    const { registry, calls } = makeFakeRegistry(withAll);

    augmentGithubCopilotRegistry(registry as never, COPILOT_TOKEN);

    expect(calls.length).toBe(0);
  });

  it('does not register when no Copilot token is available (validation would fail)', () => {
    const { registry, calls } = makeFakeRegistry([siblingModel()]);

    augmentGithubCopilotRegistry(registry as never, undefined);

    expect(calls.length).toBe(0);
    expect(registry.find('github-copilot', 'gpt-5.6-sol')).toBeUndefined();
  });

  it('derives the base URL from the token proxy-ep when no sibling model exists', () => {
    const { registry } = makeFakeRegistry([]);

    augmentGithubCopilotRegistry(registry as never, COPILOT_TOKEN);

    const found = registry.find('github-copilot', 'gpt-5.6-sol');
    expect(found).toBeDefined();
    expect(found!.baseUrl).toBe('https://api.enterprise.githubcopilot.com');
  });
});
