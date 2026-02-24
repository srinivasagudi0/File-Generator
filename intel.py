import logging
import os
import time
from pathlib import Path
import base64
from io import BytesIO
from log_utils import configure_logging, preview_text
import json

try:
    from PIL import Image, ImageChops, ImageStat
    HAS_PILLOW = True
except ImportError:
    Image = None
    ImageChops = None
    ImageStat = None
    HAS_PILLOW = False

try:
    import pytesseract
    HAS_TESSERACT = True
except Exception:
    HAS_TESSERACT = False

logger = configure_logging(__name__)

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None
    logger.warning('openai package not installed; AI features unavailable')

try:
    import requests
except ImportError:
    requests = None
    logger.warning('requests package not installed; HackClub REST fallback unavailable')

SUPPORTED_IMAGE_EXTENSIONS = ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp')
IMAGE_MIME_BY_EXT = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff',
    'webp': 'image/webp',
}
IMAGE_GENERATION_SIZES = {'1024x1024', '1024x1792', '1792x1024'}
IMAGE_EDIT_SIZES = {'256x256', '512x512', '1024x1024'}
SUPPORTED_AI_PROVIDERS = {'openai', 'hackclub'}
SUPPORTED_AI_PROVIDER_ALIASES = {'hackvcl': 'hackclub'}


def _strip_wrapped_quotes(value: str) -> str:
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def _load_local_env_file() -> None:
    candidate_paths = [
        Path.cwd() / '.env',
        Path(__file__).resolve().parent / '.env',
    ]
    seen_paths: set[Path] = set()

    for env_path in candidate_paths:
        if env_path in seen_paths:
            continue
        seen_paths.add(env_path)
        if not env_path.exists():
            continue
        try:
            loaded = 0
            for raw_line in env_path.read_text(encoding='utf-8').splitlines():
                line = raw_line.strip()
                if not line or line.startswith('#'):
                    continue
                if line.startswith('export '):
                    line = line[7:].strip()
                if '=' not in line:
                    continue
                key, value = line.split('=', 1)
                key = key.strip()
                if not key:
                    continue
                if (os.getenv(key) or '').strip():
                    continue
                os.environ[key] = _strip_wrapped_quotes(value.strip())
                loaded += 1
            logger.debug('Loaded %d variables from %s', loaded, env_path)
        except Exception as exc:
            logger.warning('Unable to load .env file from %s: %s', env_path, exc)


def _env(name: str) -> str:
    return (os.getenv(name) or '').strip()


def _env_any(*names: str) -> str:
    for name in names:
        value = _env(name)
        if value:
            return value
    return ''


def _detect_ai_provider() -> str:
    configured = _env('FILEGEN_AI_PROVIDER').lower()
    if configured in SUPPORTED_AI_PROVIDER_ALIASES:
        return SUPPORTED_AI_PROVIDER_ALIASES[configured]
    if configured in SUPPORTED_AI_PROVIDERS:
        return configured
    if _env('OPENAI_API_KEY'):
        return 'openai'
    if _env_any('HACKCLUB_API_KEY', 'HACKVCL_API_KEY'):
        return 'hackclub'
    return 'openai'


def _pick_model(
    provider: str,
    provider_keys: tuple[str, ...],
    openai_keys: tuple[str, ...],
    fallback_keys: tuple[str, ...],
    default: str,
) -> str:
    if provider == 'hackclub':
        return _env_any(*provider_keys) or _env_any(*fallback_keys) or default
    return _env_any(*openai_keys) or _env_any(*fallback_keys) or default


def _is_provider_configured(provider: str) -> bool:
    provider = (provider or '').strip().lower()
    if provider == 'hackclub':
        has_key = bool(_env_any('HACKCLUB_API_KEY', 'HACKVCL_API_KEY'))
        has_base_url = bool(_env_any('HACKCLUB_BASE_URL', 'HACKVCL_BASE_URL'))
        return has_key and has_base_url
    if provider == 'openai':
        return bool(_env('OPENAI_API_KEY'))
    return False


def _provider_retry_order() -> tuple[str, ...]:
    preferred = _detect_ai_provider()
    order: list[str] = [preferred]
    alternate = 'openai' if preferred == 'hackclub' else 'hackclub'
    if _is_provider_configured(alternate):
        order.append(alternate)
    return tuple(order)


def _load_ai_settings_for_provider(provider: str) -> dict[str, str]:
    provider = (provider or '').strip().lower()
    if provider in SUPPORTED_AI_PROVIDER_ALIASES:
        provider = SUPPORTED_AI_PROVIDER_ALIASES[provider]
    if provider not in SUPPORTED_AI_PROVIDERS:
        raise RuntimeError(f'Unsupported AI provider: {provider}')

    if provider == 'hackclub':
        api_key = _env_any('HACKCLUB_API_KEY', 'HACKVCL_API_KEY')
        if not api_key:
            raise RuntimeError('HACKCLUB_API_KEY is not set.')
        base_url = _env_any('HACKCLUB_BASE_URL', 'HACKVCL_BASE_URL')
        if not base_url:
            raise RuntimeError('HACKCLUB_BASE_URL is not set.')
    else:
        api_key = _env('OPENAI_API_KEY')
        if not api_key:
            raise RuntimeError('OPENAI_API_KEY is not set.')
        base_url = _env('OPENAI_BASE_URL')

    return {
        'provider': provider,
        'api_key': api_key,
        'base_url': base_url,
        'chat_model': _pick_model(
            provider,
            ('HACKCLUB_CHAT_MODEL', 'HACKVCL_CHAT_MODEL'),
            ('OPENAI_CHAT_MODEL',),
            ('FILEGEN_CHAT_MODEL',),
            'gpt-3.5-turbo',
        ),
        'vision_model': _pick_model(
            provider,
            ('HACKCLUB_VISION_MODEL', 'HACKVCL_VISION_MODEL'),
            ('OPENAI_VISION_MODEL',),
            ('FILEGEN_VISION_MODEL',),
            'gpt-4o-mini',
        ),
        'image_generation_model': _pick_model(
            provider,
            ('HACKCLUB_IMAGE_GENERATION_MODEL', 'HACKVCL_IMAGE_GENERATION_MODEL'),
            ('OPENAI_IMAGE_GENERATION_MODEL',),
            ('FILEGEN_IMAGE_GENERATION_MODEL',),
            'dall-e-3',
        ),
        'image_edit_models': _pick_model(
            provider,
            ('HACKCLUB_IMAGE_EDIT_MODELS', 'HACKVCL_IMAGE_EDIT_MODELS'),
            ('OPENAI_IMAGE_EDIT_MODELS',),
            ('FILEGEN_IMAGE_EDIT_MODELS',),
            'dall-e-2',
        ),
    }


def _load_ai_settings() -> dict[str, str]:
    return _load_ai_settings_for_provider(_detect_ai_provider())


def _build_ai_client(settings: dict[str, str]):
    logger.debug(
        'Building AI client | provider=%s has_key=%s has_base_url=%s',
        settings.get('provider'),
        bool(settings.get('api_key')),
        bool(settings.get('base_url')),
    )
    if OpenAI is None:
        raise RuntimeError('openai package is not installed.')

    kwargs = {
        'api_key': settings.get('api_key', ''),
        'max_retries': 2,
        'timeout': 45.0,
    }
    if settings.get('base_url'):
        kwargs['base_url'] = settings['base_url']
    return OpenAI(**kwargs)


def _parse_model_candidates(raw: str, default_model: str) -> tuple[str, ...]:
    candidates = tuple(model.strip() for model in str(raw or '').split(',') if model.strip())
    if candidates:
        return candidates
    return (default_model,)


def _chat_retry_count() -> int:
    raw = _env('FILEGEN_AI_RETRIES')
    if not raw:
        return 3
    try:
        value = int(raw)
    except ValueError:
        return 3
    return max(1, min(value, 8))


def _is_non_retryable_ai_error(exc: Exception) -> bool:
    status_code = getattr(exc, 'status_code', None)
    if isinstance(status_code, int) and status_code in (400, 401, 403, 404, 422):
        return True
    message = str(exc).lower()
    if (
        'incorrect api key' in message
        or 'invalid_api_key' in message
        or 'unsupported model' in message
        or 'model not found' in message
    ):
        return True
    return False


def _looks_like_connection_error(exc: Exception) -> bool:
    name = type(exc).__name__.lower()
    message = str(exc).lower()
    return (
        'connection' in name
        or 'timeout' in name
        or 'connection error' in message
        or 'timed out' in message
        or 'temporarily unavailable' in message
    )


def _hackclub_rest_chat_completion(
    ai_settings: dict[str, str],
    system_message: str,
    user_message: str,
) -> str:
    if requests is None:
        raise RuntimeError('requests package is not installed.')

    base_url = str(ai_settings.get('base_url') or '').rstrip('/')
    if not base_url:
        raise RuntimeError('Missing HackClub base URL.')
    endpoint = f'{base_url}/chat/completions'
    headers = {
        'Authorization': f'Bearer {ai_settings.get("api_key", "")}',
        'Content-Type': 'application/json',
    }
    payload = {
        'model': ai_settings.get('chat_model', ''),
        'messages': [
            {'role': 'system', 'content': system_message},
            {'role': 'user', 'content': user_message},
        ],
        'max_tokens': 1200,
        'temperature': 0.7,
    }
    response = requests.post(endpoint, headers=headers, json=payload, timeout=45)
    if response.status_code >= 400:
        body = response.text.strip().replace('\n', ' ')
        raise RuntimeError(f'HTTP {response.status_code}: {body[:220]}')
    try:
        data = response.json()
    except json.JSONDecodeError:
        body = response.text.strip().replace('\n', ' ')
        raise RuntimeError(f'Invalid JSON response: {body[:220]}')
    choices = data.get('choices') or []
    if not choices:
        return ''
    message = choices[0].get('message') or {}
    return str(message.get('content') or '')


def _requests_empty_content(user_input: str) -> bool:
    normalized = ' '.join(str(user_input or '').lower().split())
    if not normalized:
        return False
    phrases = (
        'write nothing',
        'nothing in it',
        'empty file',
        'blank file',
        'no content',
        'leave it empty',
    )
    return any(phrase in normalized for phrase in phrases)


_load_local_env_file()

def _get_txt_read():
    """Lazy import of extension-aware read helper from file_generator to avoid circular imports."""
    try:
        from file_generator import read
        logger.debug('Loaded read helper from file_generator successfully')
        return read
    except Exception as exc:
        logger.warning('Could not import read helper from file_generator: %s', exc)
        # return a fallback that raises when called so callers can handle it
        def _missing(name: str) -> str:
            return ''
        return _missing

def process_input(user_input: str, action:str, name: str) -> str:
    ext = Path(name).suffix.lower().lstrip('.')
    logger.info(
        'process_input invoked | action=%s file=%s ext=%s input_len=%d input_preview=%s',
        action,
        name,
        ext,
        len(str(user_input)),
        preview_text(user_input),
    )
    if ext in ('txt', 'docx', 'xlsx'):
        logger.debug('Routing process_input to stage1 for text-like extension %s', ext)
        return stage1(user_input, action, name)
    elif ext in SUPPORTED_IMAGE_EXTENSIONS:
        # pass action through to generate_image to avoid missing-argument errors
        logger.debug('Routing process_input to generate_image for extension %s', ext)
        return generate_image(user_input, name, action)
    else:
        logger.warning('Unsupported file type for %s: %s', name, ext)
        return f"Unsupported file type: {ext}. Supported: .txt, .docx, .xlsx, .png, .jpg, .jpeg, .gif, .bmp, .tiff, .webp."

def stage1(user_input: str, action:str, name: str) -> str:
    logger.info(
        'stage1 started | action=%s file=%s input_len=%d input_preview=%s',
        action,
        name,
        len(str(user_input)),
        preview_text(user_input),
    )
    reader = _get_txt_read()
    file_content = reader(name) if action in ['R', 'A'] else ''
    logger.debug(
        'stage1 loaded existing content | action=%s file=%s content_len=%d content_preview=%s',
        action,
        name,
        len(str(file_content)),
        preview_text(file_content),
    )

    provider_order = _provider_retry_order()
    logger.info('AI provider retry order | providers=%s', ','.join(provider_order))

    if action == 'W':
        message = '''You generate file-ready text with no extra commentary.
                        Only return the content that should be written into the file. 
                        If asked to create or include text in a file, just restate it clearly and concisely.
                        Do not include filenames, formats, explanations, or filler—only the file content.
                        You are allowed to put put hyperlinks in the text if relevant, but do not add any other formatting or commentary. just make it seamlessly integrate into the text.
                         You create tables always when comparing things and when needed.'''
    elif action == 'R':
        message = f'''You are to summarize the content of a file as requested by the user,
            If the input is only the file content and not a question, then you are to summarize the content of the file.
            Here is the content of the file:\n{file_content}'''
    elif action == 'A':
        message = f'''You are to add content to an existing file.
            Here is the existing content of the file:\n{file_content}
            Dont return the existing content, only the new content that should be appended to the file. 
             You create tables always when compariqing things and when needed.
             You are allowed to put put hyperlinks in the text if relevant, but do not add any other formatting or commentary. just make it seamlessly integrate into the text.
             Do not include the existing content in your response.'''
    else:
        # Usually it won't reach here without any action, but just in case
        message = f"You only say 'error in action and message type' no matter what."
    logger.debug('stage1 system message prepared | action=%s message_preview=%s', action, preview_text(message))

    if OpenAI is None:
        logger.error('AI client unavailable. Falling back to non-AI output.')
        return _fallback_output(action, user_input, file_content)

    retry_count = _chat_retry_count()
    provider_errors: list[str] = []
    for provider in provider_order:
        try:
            ai_settings = _load_ai_settings_for_provider(provider)
        except Exception as exc:
            provider_errors.append(f'{provider}: {type(exc).__name__}: {exc}')
            logger.warning('Skipping AI provider | provider=%s reason=%s', provider, exc)
            continue
        logger.info('Trying AI provider | provider=%s model=%s', provider, ai_settings.get('chat_model'))

        try:
            client = _build_ai_client(ai_settings)
        except Exception as exc:
            provider_errors.append(f'{provider}: {type(exc).__name__}: {exc}')
            logger.exception('Failed to initialize AI client | provider=%s', provider)
            continue

        for attempt in range(1, retry_count + 1):
            try:
                response = client.chat.completions.create(
                    model=ai_settings['chat_model'],
                    messages=[
                        {
                            "role": "system",
                            "content": message,
                        },
                        {"role": "user", "content": user_input},
                    ],
                    max_tokens=1200,
                    temperature=0.7,
                )
                output = response.choices[0].message.content
                if output and str(output).strip():
                    logger.info(
                        'AI completion succeeded | provider=%s model=%s action=%s file=%s output_len=%d output_preview=%s',
                        ai_settings.get('provider'),
                        ai_settings.get('chat_model'),
                        action,
                        name,
                        len(str(output)),
                        preview_text(output),
                    )
                    return output
                if action in ('W', 'A') and _requests_empty_content(user_input):
                    logger.info(
                        'AI returned empty content but request indicates blank output | action=%s file=%s',
                        action,
                        name,
                    )
                    return ''
                if attempt < retry_count:
                    logger.warning(
                        'AI returned empty content; retrying | provider=%s model=%s attempt=%d/%d',
                        ai_settings.get('provider'),
                        ai_settings.get('chat_model'),
                        attempt,
                        retry_count,
                    )
                    time.sleep(0.35 * attempt)
                    continue
                provider_errors.append(f'{provider}: empty response content')
                logger.warning(
                    'AI completion returned empty content | provider=%s model=%s action=%s file=%s',
                    ai_settings.get('provider'),
                    ai_settings.get('chat_model'),
                    action,
                    name,
                )
                break
            except Exception as exc:
                non_retryable = _is_non_retryable_ai_error(exc)
                if provider == 'hackclub' and _looks_like_connection_error(exc):
                    try:
                        logger.warning(
                            'OpenAI client connection error for HackClub; trying REST fallback | attempt=%d/%d',
                            attempt,
                            retry_count,
                        )
                        rest_output = _hackclub_rest_chat_completion(ai_settings, message, user_input)
                        if rest_output and rest_output.strip():
                            logger.info(
                                'HackClub REST fallback succeeded | model=%s action=%s file=%s output_len=%d',
                                ai_settings.get('chat_model'),
                                action,
                                name,
                                len(rest_output),
                            )
                            return rest_output
                        if action in ('W', 'A') and _requests_empty_content(user_input):
                            return ''
                        if attempt < retry_count:
                            time.sleep(0.6 * attempt)
                            continue
                        provider_errors.append(f'{provider}: empty response content (REST fallback)')
                        break
                    except Exception as rest_exc:
                        logger.warning('HackClub REST fallback failed | error=%s', rest_exc)

                if attempt < retry_count and not non_retryable:
                    logger.warning(
                        'AI completion failed; retrying | provider=%s attempt=%d/%d error=%s',
                        provider,
                        attempt,
                        retry_count,
                        exc,
                    )
                    time.sleep(0.6 * attempt)
                    continue
                provider_errors.append(f'{provider}: {type(exc).__name__}: {exc}')
                logger.exception('AI completion failed for action %s on %s using provider %s', action, name, provider)
                break

    error_message = '; '.join(provider_errors) if provider_errors else 'No AI provider is configured.'
    return _fallback_output(action, user_input, file_content, error=error_message)

def _fallback_output(action: str, user_input: str, file_content: str, error: str | None = None) -> str:
    logger.warning(
        'Using fallback output | action=%s error_present=%s user_input_len=%d file_content_len=%d',
        action,
        bool(error),
        len(str(user_input)),
        len(str(file_content)),
    )
    reason = ''
    if error:
        reason = str(error).strip().replace('\n', ' ')[:220]

    if action == 'W':
        if reason:
            return f'Error: AI unavailable. {reason}'
        return user_input
    if action == 'A':
        if reason:
            return f'Error: AI unavailable. {reason}'
        return user_input
    if action == 'R':
        if reason:
            return f'Error: AI unavailable. {reason}'
        return file_content if file_content else 'No readable content found.'
    return 'AI unavailable and no fallback available for this action.'


def _resolve_path(path_like: str) -> Path:
    path = Path(path_like).expanduser()
    logger.debug('Resolving path | raw=%s expanded=%s', path_like, path)
    if path.is_absolute():
        return path
    return (Path.cwd() / path).resolve()


def _require_pillow() -> str | None:
    if not HAS_PILLOW:
        logger.error('Pillow dependency missing')
        return 'Missing dependency: Pillow. Install with "pip install Pillow".'
    return None


def _normalize_size(size: str, allowed: set[str], default_size: str, label: str) -> str:
    normalized = (size or '').strip()
    if normalized in allowed:
        return normalized
    logger.warning('Invalid %s size "%s"; using default %s', label, size, default_size)
    return default_size

def generate_image(user_input: str, file_name: str, action:str, size: str = '1024x1024') -> str:
    action = action.upper().strip()
    logger.info(
        'generate_image invoked | action=%s file=%s size=%s prompt_len=%d prompt_preview=%s',
        action,
        file_name,
        size,
        len(str(user_input)),
        preview_text(user_input),
    )
    if action == 'W':
        return image_generation(user_input, file_name, action, size)
    elif action == 'R':
        return image_reading(file_name, user_input)
    elif action == 'A':
        return image_editing(user_input, file_name, action, size)
    else:
        return "Invalid action for image processing."

def image_generation(user_input: str, file_name: str, action:str, size: str = '1024x1024'):
    """Generates an image using the HackClub (OpenAI-compatible) images API and saves it to `output_path`.
                Raises RuntimeError on missing key, API failure, decode error, or file write error.
                Returns the path to the saved file."""

    prompt = (user_input or '').strip()
    logger.debug('image_generation prompt normalized | len=%d preview=%s', len(prompt), preview_text(prompt))
    if not prompt:
        return 'Error: image generation prompt cannot be empty.'

    normalized_size = _normalize_size(size, IMAGE_GENERATION_SIZES, '1024x1024', 'image generation')

    try:
        ai_settings = _load_ai_settings()
        client = _build_ai_client(ai_settings)
    except Exception as exc:
        logger.exception('Failed to initialize AI client for image generation')
        return f'Error initializing AI client: {exc}'

    try:
        response = client.images.generate(
            model=ai_settings['image_generation_model'],
            prompt=prompt,
            size=normalized_size,
            n=1,
            response_format="b64_json",
        )
    except Exception as e:
        logger.exception('Image generation request failed for %s', file_name)
        return f'Error generating image: {e}'
    try:
        if not getattr(response, 'data', None):
            return 'Error: Empty response data during image generation.'
        item = response.data[0]
        image_b64 = item.get("b64_json") if isinstance(item, dict) else getattr(item, "b64_json", None)
        if not image_b64:
            return 'Error: No image data returned during image generation.'
        image_bytes = base64.b64decode(image_b64, validate=True)
        logger.debug('image_generation decoded bytes | file=%s bytes=%d', file_name, len(image_bytes))
        out_path = _resolve_path(file_name)
        out_path = out_path.with_suffix('.png')
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(image_bytes)
        logger.info('Generated image saved to %s', out_path)
        return str(out_path.resolve())
    except Exception as e:
        logger.exception('Failed to decode/save generated image for %s', file_name)
        return f'Error saving generated image: {e}'


def _prep_image_for_edit(img_path: Path, target_size: int = 1024) -> tuple[BytesIO, BytesIO]:
    logger.debug('Preparing image for edit | path=%s target_size=%d', img_path, target_size)
    with Image.open(img_path) as img:
        rgba = img.convert("RGBA")
        rgba.thumbnail((target_size, target_size), Image.LANCZOS)
        canvas = Image.new("RGBA", (target_size, target_size), (255, 255, 255, 0))
        x = (target_size - rgba.width) // 2
        y = (target_size - rgba.height) // 2
        canvas.paste(rgba, (x, y))

        # Image edit masks use transparency for editable regions.
        # Fully transparent mask => whole image is editable.
        mask = Image.new("RGBA", (target_size, target_size), (0, 0, 0, 0))

        buf_img = BytesIO()
        canvas.save(buf_img, format="PNG")
        buf_img.seek(0)
        buf_img.name = "image.png"

        buf_mask = BytesIO()
        mask.save(buf_mask, format="PNG")
        buf_mask.seek(0)
        buf_mask.name = "mask.png"
        return buf_img, buf_mask


def _image_diff_ratio(original_path: Path, edited_bytes: bytes) -> float:
    logger.debug('Computing image diff ratio | original=%s edited_bytes=%d', original_path, len(edited_bytes))
    with Image.open(original_path) as original:
        original_rgb = original.convert("RGB")
        with Image.open(BytesIO(edited_bytes)) as edited:
            edited_rgb = edited.convert("RGB").resize(original_rgb.size, Image.LANCZOS)
        diff = ImageChops.difference(original_rgb, edited_rgb)
        stat = ImageStat.Stat(diff)
    # Normalize mean absolute difference across RGB channels to 0..1
    return sum(stat.mean) / (len(stat.mean) * 255.0)


def _is_grayscale_request(user_input: str) -> bool:
    lowered = user_input.lower()
    logger.debug('Checking grayscale request | preview=%s', preview_text(user_input))
    return any(
        kw in lowered
        for kw in ('black and white', 'grayscale', 'grey scale', 'greyscale', 'gray scale', 'monochrome')
    )


def _save_processed_image(image_bytes: bytes, source_path: Path, suffix: str = "_edited") -> str:
    out_dir = Path('processed_images')
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f'{source_path.stem}{suffix}.png'
    out_path.write_bytes(image_bytes)
    logger.info('Processed image saved to %s', out_path)
    return str(out_path.resolve())


def _apply_local_grayscale(image_path: Path) -> str:
    logger.info('Applying local grayscale conversion | file=%s', image_path)
    with Image.open(image_path) as img:
        img_bw = img.convert('L').convert('RGBA')
        buf = BytesIO()
        img_bw.save(buf, format='PNG')
        return _save_processed_image(buf.getvalue(), image_path, suffix='_bw')


def _request_image_edit(
    client: OpenAI,
    img_buf: BytesIO,
    mask_buf: BytesIO,
    prompt: str,
    size: str,
    model_candidates: tuple[str, ...],
):
    last_exc = None
    logger.debug('Requesting image edit | size=%s prompt_preview=%s models=%s', size, preview_text(prompt), model_candidates)
    for model_name in model_candidates:
        try:
            img_buf.seek(0)
            mask_buf.seek(0)
            response = client.images.edit(
                model=model_name,
                image=img_buf,
                mask=mask_buf,
                prompt=prompt,
                size=size,
                n=1,
                response_format="b64_json",
            )
            logger.info('Image edit succeeded with model %s', model_name)
            return response
        except Exception as exc:
            last_exc = exc
            logger.warning('Image edit failed with model %s: %s', model_name, exc)
    raise RuntimeError(f'Image edit failed for all models: {last_exc}')


def _decode_image_response(response) -> bytes:
    logger.debug('Decoding image response payload')
    if not getattr(response, 'data', None):
        raise RuntimeError('Empty response data during image edit.')
    item = response.data[0]
    image_b64 = item.get("b64_json") if isinstance(item, dict) else getattr(item, "b64_json", None)
    if not image_b64:
        raise RuntimeError('No image data returned during edit.')
    return base64.b64decode(image_b64, validate=True)


def image_editing(user_input: str, file_name: str, action: str, size: str = '1024x1024') -> str:
    # Edits an existing image based on the user instruction and saves to processed_images/
    logger.info(
        'image_editing invoked | action=%s file=%s size=%s prompt_len=%d prompt_preview=%s',
        action,
        file_name,
        size,
        len(str(user_input)),
        preview_text(user_input),
    )
    missing_dep = _require_pillow()
    if missing_dep:
        logger.error(missing_dep)
        return missing_dep

    prompt = (user_input or '').strip()
    if not prompt:
        return 'Error: image edit prompt cannot be empty.'

    normalized_size = _normalize_size(size, IMAGE_EDIT_SIZES, '1024x1024', 'image editing')

    image_path = _resolve_path(file_name)
    if not image_path.exists():
        return f'Error: image file not found: {file_name}'

    # Deterministic local path for black-and-white requests (avoids no-op API edits)
    if _is_grayscale_request(prompt):
        try:
            return _apply_local_grayscale(image_path)
        except Exception as exc:
            logger.exception('Local grayscale processing failed for %s', file_name)
            return f'Error applying grayscale edit: {exc}'

    try:
        ai_settings = _load_ai_settings()
        client = _build_ai_client(ai_settings)
        edit_models = _parse_model_candidates(ai_settings.get('image_edit_models', ''), 'dall-e-2')
    except Exception as exc:
        logger.exception('Failed to initialize AI client for image editing')
        return f'Error initializing AI client: {exc}'

    try:
        img_buf, mask_buf = _prep_image_for_edit(image_path)

        strong_prompt = (
            prompt
            + ". Make the requested edits clearly visible—do not just change brightness or contrast. "
              "If removal is requested, ensure the subject is completely absent. "
              "Apply high-strength changes and honor all described alterations."
        )
        response = _request_image_edit(client, img_buf, mask_buf, strong_prompt, normalized_size, edit_models)
    except Exception as exc:
        logger.exception('Image edit request failed for %s', file_name)
        return f'Error editing image: {exc}'

    try:
        image_bytes = _decode_image_response(response)
        diff_ratio = _image_diff_ratio(image_path, image_bytes)
        logger.info('Image edit diff ratio for %s: %.4f', file_name, diff_ratio)

        # Retry once with a stricter prompt if output is too similar.
        if diff_ratio < 0.02:
            logger.warning('Edited output too similar to original; retrying with stricter prompt for %s', file_name)
            img_buf, mask_buf = _prep_image_for_edit(image_path)
            retry_prompt = (
                strong_prompt
                + " The output must be visibly different from the input and clearly apply the requested edit."
            )
            retry_response = _request_image_edit(client, img_buf, mask_buf, retry_prompt, normalized_size, edit_models)
            image_bytes = _decode_image_response(retry_response)
            diff_ratio = _image_diff_ratio(image_path, image_bytes)
            logger.info('Image edit retry diff ratio for %s: %.4f', file_name, diff_ratio)

        if diff_ratio < 0.01:
            logger.error('Edited output too similar to original for %s after retry', file_name)
            return 'Error: edited image was too similar to the original. Try a more specific edit prompt.'

        return _save_processed_image(image_bytes, image_path)
    except Exception as exc:
        logger.exception('Failed to decode/save edited image for %s', file_name)
        return f'Error saving edited image: {exc}'

def image_reading(path: str, user_request: str = '') -> str:
    logger.info(
        'image_reading invoked | path=%s user_request_len=%d user_request_preview=%s',
        path,
        len(str(user_request)),
        preview_text(user_request),
    )
    missing_dep = _require_pillow()
    if missing_dep:
        logger.error(missing_dep)
        return missing_dep

    image_path = _resolve_path(path)
    ext = image_path.suffix.lower().lstrip('.')
    if ext not in SUPPORTED_IMAGE_EXTENSIONS:
        return f'Error: unsupported image type ".{ext}".'
    if not image_path.exists():
        return f'Error: image file not found: {path}'

    def ocr_extract(img_path: Path) -> str:
        if not HAS_TESSERACT:
            logger.warning('Tesseract not available; skipping OCR for %s', img_path)
            return ""
        try:
            with Image.open(img_path) as img:
                text = pytesseract.image_to_string(img).strip()
                logger.info('OCR extraction completed | file=%s text_len=%d', img_path, len(text))
                return text
        except Exception as exc:
            logger.warning('OCR failed for %s: %s', img_path, exc)
            return ""

    def image_info(img_path: Path) -> str:
        try:
            with Image.open(img_path) as img:
                width, height = img.size
                mode = img.mode
                fmt = img.format
            return f'Image info: format={fmt}, size={width}x{height}, mode={mode}.'
        except Exception as exc:
            logger.warning('Failed to read image metadata for %s: %s', img_path, exc)
            return 'Image info unavailable.'

    def extract_response_text(resp) -> str:
        out = getattr(resp, "output_text", None)
        if out:
            return out.strip()
        parts = []
        for item in getattr(resp, "output", []) or []:
            for c in item.get("content", []) or []:
                if c.get("type") == "output_text":
                    parts.append(c.get("text", ""))
        return ("\n".join(parts) or str(resp)).strip()

    def summarize_text(text: str, client: OpenAI, model: str) -> str:
        focus = user_request.strip()
        if focus:
            prompt = (
                "The user asked for this focus: " + focus + "\n\n"
                "Analyze OCR text from an image. Give a focused answer in bullets and a short summary.\n\n"
                + text
            )
        else:
            prompt = "Summarize the text below in 3 bullets and one-line summary:\n\n" + text
        resp = client.responses.create(model=model, input=prompt)
        logger.info('Text summarization via Responses API completed | model=%s input_len=%d', model, len(text))
        return extract_response_text(resp)

    def summarize_image_via_responses(img_path: Path, client: OpenAI, model: str) -> str:
        mime = IMAGE_MIME_BY_EXT.get(ext, 'image/png')
        with open(img_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        data_url = f"data:{mime};base64,{b64}"
        multimodal_input = [
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": (
                            "Describe this image in detail and then provide a concise summary."
                            if not user_request.strip()
                            else f"User focus: {user_request.strip()}. Describe only what helps answer that request, then summarize."
                        ),
                    },
                    {"type": "input_image", "image_url": data_url}
                ]
            }
        ]
        resp = client.responses.create(model=model, input=multimodal_input, max_output_tokens=500)
        logger.info('Image summarization via multimodal Responses API completed | model=%s file=%s', model, img_path)
        return extract_response_text(resp)

    # Attempt OCR first; fall back to multimodal if needed.
    text_from_image = ocr_extract(image_path)
    info = image_info(image_path)

    if OpenAI is None:
        logger.error('AI client unavailable for image reading.')
        focus_line = f'User focus: {user_request.strip()}\n\n' if user_request.strip() else ''
        if text_from_image:
            return f'{focus_line}{info}\n\nOCR text:\n{text_from_image}'
        return f'{focus_line}{info}\n\nAI client unavailable and no OCR text extracted.'

    try:
        ai_settings = _load_ai_settings()
        client = _build_ai_client(ai_settings)
    except Exception as exc:
        logger.exception('Failed to initialize AI client for image reading')
        focus_line = f'User focus: {user_request.strip()}\n\n' if user_request.strip() else ''
        if text_from_image:
            return f'{focus_line}{info}\n\nOCR text:\n{text_from_image}'
        return f'{focus_line}Error initializing AI client: {exc}'

    try:
        if len(text_from_image) >= 40:
            return summarize_text(text_from_image, client, ai_settings['vision_model'])
        return summarize_image_via_responses(image_path, client, ai_settings['vision_model'])
    except Exception as exc:
        logger.exception('Image summarization failed for %s', image_path)
        focus_line = f'User focus: {user_request.strip()}\n\n' if user_request.strip() else ''
        if text_from_image:
            return f'{focus_line}{info}\n\nOCR text:\n{text_from_image}'
        return f'{focus_line}{info}\n\nError summarizing image: {exc}'


# End of intel.py EOF
if __name__ == '__main__':
    print('End of intel.py EOF and no errors found')
