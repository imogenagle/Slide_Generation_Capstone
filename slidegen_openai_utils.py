import os
from typing import List, Optional

from openai import AzureOpenAI, OpenAI

from camel.types import ModelPlatformType


DEFAULT_AZURE_OPENAI_BASE_URL = "https://slidegen-pers-openai.openai.azure.com/"
DEFAULT_AZURE_API_VERSION = "2024-10-21"
DEFAULT_AZURE_DEPLOYMENT_NAME = "gpt-4o"


def apply_default_azure_openai_env() -> None:
    os.environ.setdefault("AZURE_OPENAI_BASE_URL", DEFAULT_AZURE_OPENAI_BASE_URL)
    os.environ.setdefault("AZURE_API_VERSION", DEFAULT_AZURE_API_VERSION)
    os.environ.setdefault("AZURE_DEPLOYMENT_NAME", DEFAULT_AZURE_DEPLOYMENT_NAME)


def _normalize_model_name(model_name: Optional[str]) -> str:
    return (model_name or "").strip().lower().replace("_", "-")


def azure_openai_enabled() -> bool:
    apply_default_azure_openai_env()

    explicit = os.getenv("SLIDEGEN_USE_AZURE_OPENAI", "").strip().lower()
    if explicit in {"1", "true", "yes", "on"}:
        return True
    if explicit in {"0", "false", "no", "off"}:
        return False

    return bool(
        os.getenv("AZURE_OPENAI_API_KEY")
        and os.getenv("AZURE_OPENAI_BASE_URL")
        and os.getenv("AZURE_DEPLOYMENT_NAME")
        and os.getenv("AZURE_API_VERSION")
    )


def _azure_aliases_for_deployment(deployment_name: Optional[str] = None) -> set[str]:
    deployment = _normalize_model_name(
        deployment_name or os.getenv("AZURE_DEPLOYMENT_NAME")
    )
    aliases = {deployment}

    alias_map = {
        "gpt-4o": {"4o", "gpt-4o", "gpt4o", "gpt-4o-2024-08-06"},
        "gpt-4o-mini": {
            "4o-mini",
            "gpt-4o-mini",
            "gpt4omini",
            "gpt-4o-mini-2024-07-18",
        },
        "gpt-4.1": {"gpt-4.1"},
        "gpt-4.1-mini": {"gpt-4.1-mini"},
        "o1": {"o1"},
        "o3": {"o3"},
        "o3-mini": {"o3-mini"},
        "gpt-5": {"gpt-5", "gpt5"},
        "gpt-5.4-nano": {"gpt-5", "gpt5", "gpt-5.4-nano"},
    }
    aliases.update(alias_map.get(deployment, set()))
    return aliases


def azure_openai_supports_model(model_name: Optional[str]) -> bool:
    return azure_openai_enabled() and (
        _normalize_model_name(model_name) in _azure_aliases_for_deployment()
    )


def azure_display_models() -> List[str]:
    if not azure_openai_enabled():
        return []

    deployment = _normalize_model_name(os.getenv("AZURE_DEPLOYMENT_NAME"))
    preferred = {
        "gpt-4o": ["4o"],
        "gpt-4o-mini": ["4o-mini"],
        "gpt-4.1": ["gpt-4.1"],
        "gpt-4.1-mini": ["gpt-4.1-mini"],
        "o1": ["o1"],
        "o3": ["o3"],
        "o3-mini": ["o3-mini"],
        "gpt-5": ["gpt-5"],
        "gpt-5.4-nano": ["gpt-5"],
    }
    return preferred.get(deployment, [os.getenv("AZURE_DEPLOYMENT_NAME", deployment)])


def resolve_model_platform(
    model_name: Optional[str],
    default_platform: ModelPlatformType,
) -> ModelPlatformType:
    if (
        default_platform == ModelPlatformType.OPENAI
        and azure_openai_supports_model(model_name)
    ):
        return ModelPlatformType.AZURE
    return default_platform


def resolve_direct_model_name(model_name: str) -> str:
    if azure_openai_supports_model(model_name):
        return os.getenv("AZURE_DEPLOYMENT_NAME", model_name)
    return model_name


def should_use_direct_openai_client(model_name: Optional[str]) -> bool:
    normalized = _normalize_model_name(model_name)
    return "gpt-5" in normalized or azure_openai_supports_model(model_name)


def build_openai_client(base_url: Optional[str] = None, api_key: Optional[str] = None):
    if base_url:
        return OpenAI(base_url=base_url, api_key=api_key)

    if azure_openai_enabled():
        return AzureOpenAI(
            azure_endpoint=os.environ["AZURE_OPENAI_BASE_URL"],
            api_version=os.environ["AZURE_API_VERSION"],
            api_key=api_key or os.getenv("AZURE_OPENAI_API_KEY"),
        )

    return OpenAI(base_url=base_url, api_key=api_key)
