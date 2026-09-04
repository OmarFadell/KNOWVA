import config from "../../config";
import { AnthropicProvider } from "./anthropic";
import type { LlmProvider } from "./provider";

export * from "./provider";

/**
 * Selects the provider named by LLM_PROVIDER.
 *
 * Adding a provider is a new file implementing LlmProvider plus a case here.
 * Callers depend only on the interface, so nothing else changes.
 */
export function createLlmProvider(): LlmProvider {
  switch (config.llm.provider) {
    case "anthropic":
      return new AnthropicProvider({
        apiKey: config.llm.anthropicApiKey,
        model: config.llm.anthropicModel,
      });
    default:
      throw new Error(
        `Unsupported LLM_PROVIDER "${config.llm.provider}". Supported values: anthropic.`
      );
  }
}
