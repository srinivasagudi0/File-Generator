from __future__ import annotations

import mimetypes
import re
from pathlib import Path
from uuid import uuid4

import streamlit as st

from file_generator import agent
from intel import process_input as process
from recipe_support import (
    ACTION_LABELS as RECIPE_ACTION_LABELS,
    append_warning_for_item,
    build_recipe_document,
    build_document_from_cards,
    build_pack_from_template,
    default_file_pack_card,
    dump_recipe_document,
    get_starter_template_choices,
    hydrate_cards_from_document,
    load_recipe_text,
    preview_file_pack_cards,
    recipe_sample_text,
    summarize_recipe_item,
)

IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp'}
PROCESS_SUPPORTED_EXTENSIONS = {'txt', 'docx', 'xlsx', 'xls'}
TEXT_PREVIEW_EXTENSIONS = {
    'txt', 'md', 'markdown', 'html', 'htm', 'css',
    'py', 'js', 'java', 'c', 'cpp', 'rb', 'go', 'rs', 'swift', 'kt', 'ts', 'php',
    'json', 'xml', 'yml', 'yaml', 'toml', 'ini', 'sh', 'bat', 'ps1', 'csv', 'svg',
}
DEFAULT_EXT_BY_TYPE = {
    'txt': 'txt',
    'docx': 'docx',
    'xlsx': 'xlsx',
    'csv': 'csv',
    'pdf': 'pdf',
    'pptx': 'pptx',
    'markdown': 'md',
    'html': 'html',
    'code': 'py',
    'image': 'png',
    'chart': 'png',
    'audio': 'mp3',
    'video': 'mp4',
}
SUPPORTED_FILE_TYPES = [
    'txt', 'docx', 'xlsx', 'csv', 'pdf', 'pptx', 'markdown',
    'html', 'code', 'image', 'chart', 'audio', 'video',
]
CREATE_ACTION_LABELS = {'Write': 'W', 'Append': 'A'}
READ_DELETE_ACTION_LABELS = {'Read': 'R', 'Delete': 'D'}
FILE_TYPE_HINTS = {
    'txt': 'Summaries, instructions, quick notes, or templates.',
    'docx': 'Reports, briefs, meeting notes with headings and bullet lists.',
    'xlsx': 'Budgets, trackers, comparison tables with sheet names.',
    'csv': 'Data exports; mention headers and delimiters.',
    'pdf': 'Polished docs; specify margins, sections, or cover page.',
    'pptx': 'Slides; list slide titles, bullets, and speaker notes.',
    'markdown': 'Docs for GitHub/Wikis; headings, code blocks, and links.',
    'html': 'Landing pages or emails; include hero text, CTA, sections.',
    'code': 'Provide language, purpose, function signatures, and comments.',
    'image': 'Describe subject, style, lighting, composition.',
    'chart': 'Choose chart type and supply label:value pairs.',
    'audio': 'Voiceover scripts or podcast outline.',
    'video': 'Storyboard beats, scenes, or transcript.',
}
SUPPORTED_DETAIL_CATEGORIES = (
    'images', 'tables', 'charts', 'graphs', 'hyperlinks',
    'fonts', 'colors', 'sizes', 'styles', 'alignments',
    'margins', 'paddings', 'borders', 'backgrounds', 'layouts',
    'templates', 'themes', 'sections', 'headers', 'footers',
    'page_numbers', 'tables_of_contents', 'indexes', 'bibliographies',
    'citations', 'footnotes', 'notes',
)
DRAFT_DIR = Path('.ui_previews')


def parse_chart_data(raw_data: str) -> tuple[dict[str, float], str]:
    if not raw_data.strip():
        return {}, 'Chart data cannot be empty.'
    entries = [entry.strip() for entry in re.split(r'[,;\n]+', raw_data) if entry.strip()]
    parsed: dict[str, float] = {}
    seen: set[str] = set()
    for entry in entries:
        if ':' not in entry:
            return {}, f'Invalid chart entry "{entry}". Use label:value.'
        label, value = entry.split(':', 1)
        label = label.strip()
        if not label:
            return {}, 'Chart label cannot be empty.'
        key = label.casefold()
        if key in seen:
            return {}, f'Duplicate chart label "{label}" is not allowed.'
        try:
            parsed[label] = float(value.strip().replace('%', ''))
        except ValueError:
            return {}, f'Chart value for "{label}" must be numeric.'
        seen.add(key)
    return parsed, ''


def parse_details(raw: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if ':' in line:
            category, value = line.split(':', 1)
        elif '=' in line:
            category, value = line.split('=', 1)
        else:
            items.append(('notes', line))
            continue
        category = re.sub(r'[\s\-]+', '_', category.strip().lower())
        value = value.strip()
        if not value:
            continue
        if category not in SUPPORTED_DETAIL_CATEGORIES:
            items.append(('notes', f'{category}: {value}'))
            continue
        items.append((category, value))
    return items


def _validate_generation_inputs(
    create_action: str,
    file_type: str,
    content: str,
    chart_type: str,
    chart_data_raw: str,
    append_target: str,
) -> tuple[list[str], list[str]]:
    """Return (errors, warnings) for the creation flow."""
    errors: list[str] = []
    warnings: list[str] = []

    if file_type != 'chart' and not content.strip():
        errors.append('Content is required for this file type.')

    if file_type == 'chart':
        _, chart_error = parse_chart_data(chart_data_raw)
        if chart_error:
            errors.append(chart_error)
        if not chart_type:
            errors.append('Choose a chart type.')

    if create_action == 'A' and append_target.strip():
        target = append_target.strip()
        target_path = Path(target if '.' in Path(target).name else resolve_file_name(target, file_type))
        if not target_path.exists():
            warnings.append(f'Append target "{target_path.name}" was not found. A new draft will be created instead.')

    return errors, warnings


def _show_step_feedback(errors: list[str], warnings: list[str], ready_text: str) -> None:
    """Render inline feedback for the current step."""
    if errors:
        bullets = '\n'.join(f'- {msg}' for msg in errors)
        st.error(f'Please fix these before continuing:\n{bullets}')
        return
    if warnings:
        bullets = '\n'.join(f'- {msg}' for msg in warnings)
        st.warning(f'Heads up:\n{bullets}')
    st.success(ready_text)


def resolve_file_name(name: str, file_type: str) -> str:
    cleaned = (name or '').strip()
    if not cleaned:
        return f'output.{DEFAULT_EXT_BY_TYPE[file_type]}'
    if '.' not in cleaned:
        return f'{cleaned}.{DEFAULT_EXT_BY_TYPE[file_type]}'
    return cleaned


def supports_ai_processing(ext: str) -> bool:
    return ext in PROCESS_SUPPORTED_EXTENSIONS


def _collect_format_options(
    font: str,
    color: str,
    size: float,
    styles_raw: str,
    alignment: str,
) -> dict[str, object]:
    format_options: dict[str, object] = {}
    if font.strip():
        format_options['font'] = font.strip()
    if color.strip():
        format_options['color'] = color.strip()
    if size > 0:
        format_options['size'] = float(size)
    style_tokens = [token.strip().lower() for token in styles_raw.split(',') if token.strip()]
    if style_tokens:
        format_options['styles'] = style_tokens
    if alignment:
        format_options['alignment'] = alignment
    return format_options


def run_action(
    action: str,
    file_name: str,
    content: str,
    chart_type: str,
    chart_data_raw: str,
    style: str,
    format_options: dict[str, object],
    detail_items: list[tuple[str, str]],
) -> str:
    ext = Path(file_name).suffix.lower().lstrip('.')
    is_image = ext in IMAGE_EXTENSIONS

    if action == 'D':
        return str(agent(content, file_name, action, style, format_options=format_options))

    if action == 'R':
        if is_image:
            return str(process(content, action, file_name))
        read_result = agent('', file_name, action, style, format_options=format_options)
        if supports_ai_processing(ext):
            return str(process(content, action, file_name))
        return str(read_result)

    if action in ('W', 'A') and chart_type:
        chart_data, error = parse_chart_data(chart_data_raw)
        if error:
            return f'Error: {error}'
        if not chart_data:
            return 'Error: Chart data cannot be empty.'
        result = agent('', file_name, action, style, chart_type=chart_type, chart_data=chart_data)
        return str(result)

    if is_image:
        payload = content
    elif supports_ai_processing(ext):
        payload = process(content, action, file_name)
        if _is_error_text(payload):
            return str(payload)
        if isinstance(payload, str) and payload.lstrip().lower().startswith('[ai unavailable'):
            return payload
        if action in ('W', 'A') and content.strip() and not str(payload).strip():
            return 'AI returned empty content. Please check your AI API settings and try again.'
    else:
        payload = content
    result = agent(payload, file_name, action, style, format_options=format_options, details=detail_items)
    return str(result)


def _build_recipe_item_from_form(
    action: str,
    file_type: str,
    file_name: str,
    content: str,
    chart_type: str,
    chart_data_raw: str,
    style: str,
    format_options: dict[str, object],
    detail_items: list[tuple[str, str]],
) -> dict[str, object]:
    return {
        'action': action,
        'file_type': file_type,
        'file_name': file_name,
        'content': content,
        'style': style.strip(),
        'format_options': dict(format_options),
        'details': list(detail_items),
        'chart_type': chart_type if file_type == 'chart' else '',
        'chart_data_raw': chart_data_raw if file_type == 'chart' else '',
    }


def run_recipe_batch(
    items: list[dict[str, object]],
    progress_callback=None,
) -> dict[str, object]:
    results: list[dict[str, object]] = []
    success_count = 0

    for index, item in enumerate(items, start=1):
        result = run_action(
            action=str(item.get('action', 'W')),
            file_name=str(item.get('file_name', '')),
            content=str(item.get('content', '')),
            chart_type=str(item.get('chart_type', '')),
            chart_data_raw=str(item.get('chart_data_raw', '')),
            style=str(item.get('style', '')),
            format_options=dict(item.get('format_options', {})),
            detail_items=list(item.get('details', [])),
        )
        success = not _is_error_text(result)
        saved_path = ''
        if success:
            result_path = Path(str(result))
            target_path = Path(str(item.get('file_name', '')))
            if result_path.exists():
                saved_path = str(result_path.resolve())
            elif target_path.exists():
                saved_path = str(target_path.resolve())
            success_count += 1

        recipe_result = {
            'index': index,
            'action': str(item.get('action', 'W')),
            'file_type': str(item.get('file_type', '')),
            'file_name': str(item.get('file_name', '')),
            'success': success,
            'message': str(result),
            'saved_path': saved_path,
        }
        results.append(recipe_result)
        if progress_callback is not None:
            progress_callback(index, len(items), recipe_result)

    return {
        'total': len(items),
        'success_count': success_count,
        'failure_count': len(items) - success_count,
        'results': results,
    }


def _init_state() -> None:
    if 'draft' not in st.session_state:
        st.session_state.draft = None
    if 'feedback_text' not in st.session_state:
        st.session_state.feedback_text = ''
    if 'feedback_history' not in st.session_state:
        st.session_state.feedback_history = []
    if 'feedback_notice' not in st.session_state:
        st.session_state.feedback_notice = None
    if 'recipe_document' not in st.session_state:
        st.session_state.recipe_document = None
    if 'recipe_errors' not in st.session_state:
        st.session_state.recipe_errors = []
    if 'recipe_source_name' not in st.session_state:
        st.session_state.recipe_source_name = ''
    if 'recipe_results' not in st.session_state:
        st.session_state.recipe_results = None
    if 'file_pack_name' not in st.session_state:
        st.session_state.file_pack_name = 'My File Pack'
    if 'file_pack_cards' not in st.session_state:
        st.session_state.file_pack_cards = [default_file_pack_card()]
    if 'file_pack_template' not in st.session_state:
        st.session_state.file_pack_template = 'blank'
    if 'file_pack_notes' not in st.session_state:
        st.session_state.file_pack_notes = []


def _new_draft_path(file_type: str) -> str:
    DRAFT_DIR.mkdir(parents=True, exist_ok=True)
    ext = DEFAULT_EXT_BY_TYPE[file_type]
    return str((DRAFT_DIR / f'draft_{uuid4().hex}.{ext}').resolve())


def _guess_mime(path: Path) -> str:
    guessed, _ = mimetypes.guess_type(str(path))
    if guessed:
        return guessed
    return 'application/octet-stream'


def _render_preview(path: Path, file_type: str) -> None:
    st.subheader('Preview')
    st.caption(f'Draft file: `{path}`')
    ext = path.suffix.lower().lstrip('.')

    if ext in IMAGE_EXTENSIONS:
        st.image(str(path), use_container_width=True)
        return

    if ext in TEXT_PREVIEW_EXTENSIONS:
        try:
            text = path.read_text(encoding='utf-8', errors='replace')
        except Exception as exc:
            st.error(f'Unable to read text preview: {exc}')
            return
        lowered_text = text.lstrip().lower()
        if lowered_text.startswith('[ai unavailable') or lowered_text.startswith('error: ai unavailable'):
            st.error('AI generation failed for this draft. Update your AI API settings and generate a new preview.')
        st.text_area('File Preview', value=text, height=320, disabled=True)
        return

    summary = agent('', str(path), 'R')
    st.text_area('File Preview', value=str(summary), height=320, disabled=True)


def _ai_rewrite_from_feedback(current_content: str, feedback: str) -> str:
    prompt = (
        'Rewrite the file content by applying the requested changes.\n'
        'Return only the final revised file content.\n\n'
        f'Current content:\n{current_content}\n\n'
        f'Requested changes:\n{feedback}\n'
    )
    return str(process(prompt, 'W', 'ui_feedback.txt'))


def _parse_chart_feedback(feedback: str, default_type: str) -> tuple[str, str, str]:
    raw = str(feedback or '').strip()
    if not raw:
        return '', '', 'Feedback cannot be empty.'
    match = re.match(r'^\s*(line|bar|pie|scatter)\s*[\|:]\s*(.+)$', raw, flags=re.IGNORECASE)
    if match:
        return match.group(1).lower(), match.group(2).strip(), ''
    if default_type:
        return default_type, raw, ''
    return '', '', 'For chart updates, use format like "line|Jan:10,Feb:20".'


def _apply_feedback_to_draft(
    feedback: str,
    style: str,
    format_options: dict[str, object],
    detail_items: list[tuple[str, str]],
) -> tuple[bool, str]:
    draft = st.session_state.draft
    if not draft:
        return False, 'No draft available yet.'
    path = Path(draft['path'])
    if not path.exists():
        return False, 'Draft file not found. Generate a preview again.'

    file_type = draft['file_type']
    ext = path.suffix.lower().lstrip('.')

    if file_type == 'image' or ext in IMAGE_EXTENSIONS:
        result = agent(feedback, str(path), 'A')
        return (not _is_error_text(result), str(result))

    if file_type == 'chart':
        chart_type, chart_data_raw, err = _parse_chart_feedback(feedback, draft.get('chart_type', ''))
        if err:
            return False, err
        chart_data, parse_error = parse_chart_data(chart_data_raw)
        if parse_error:
            return False, parse_error
        result = agent('', str(path), 'W', chart_type=chart_type, chart_data=chart_data)
        if not _is_error_text(result):
            draft['chart_type'] = chart_type
        return (not _is_error_text(result), str(result))

    if file_type in ('audio', 'video'):
        return False, 'Feedback-based editing is not supported for audio/video yet. Use a new source path.'

    current_content = str(agent('', str(path), 'R'))
    lowered_content = current_content.lstrip().lower()
    if lowered_content.startswith('[ai unavailable') or lowered_content.startswith('error: ai unavailable'):
        return False, 'Current draft was created in fallback mode. Generate a new preview first.'
    rewritten = _ai_rewrite_from_feedback(current_content, feedback)
    if _is_error_text(rewritten):
        return False, rewritten

    result = agent(
        rewritten,
        str(path),
        'W',
        style,
        format_options=format_options,
        details=detail_items,
    )
    return (not _is_error_text(result), str(result))


def _set_feedback_notice(level: str, message: str) -> None:
    st.session_state.feedback_notice = {'level': level, 'message': message}


def _render_feedback_notice() -> None:
    notice = st.session_state.feedback_notice
    if not notice:
        return
    level = str(notice.get('level', 'info'))
    message = str(notice.get('message', ''))
    if message:
        if level == 'success':
            st.success(message)
        elif level == 'warning':
            st.warning(message)
        elif level == 'error':
            st.error(message)
        else:
            st.info(message)
    st.session_state.feedback_notice = None


def _handle_apply_changes() -> None:
    feedback = str(st.session_state.get('feedback_text', '')).strip()
    if not feedback:
        _set_feedback_notice('warning', 'Enter change instructions first.')
        return

    draft = st.session_state.draft or {}
    success, message = _apply_feedback_to_draft(
        feedback=feedback,
        style=str(draft.get('style', '')),
        format_options=dict(draft.get('format_options', {})),
        detail_items=list(draft.get('detail_items', [])),
    )
    st.session_state.feedback_history.append(feedback)
    if success:
        st.session_state.feedback_text = ''
        _set_feedback_notice('success', 'Draft updated. Review the new preview.')
    else:
        _set_feedback_notice('error', message)


def _is_error_text(value: object) -> bool:
    if not isinstance(value, str):
        return False
    lowered = value.lower()
    return (
        lowered.startswith('error')
        or lowered.startswith('[ai unavailable')
        or 'ai unavailable' in lowered
        or 'missing dependency' in lowered
        or 'unsupported file type' in lowered
        or 'invalid action' in lowered
        or 'not found' in lowered
        or 'ai returned empty content' in lowered
    )


def _render_sidebar_help() -> None:
    draft = st.session_state.draft
    st.sidebar.header('Navigation')
    st.sidebar.markdown('1. Create a draft preview')
    st.sidebar.markdown('2. Review the preview')
    st.sidebar.markdown('3. Request changes')
    st.sidebar.markdown('4. Save only when satisfied')
    if draft and Path(draft['path']).exists():
        st.sidebar.success('Draft ready')
        st.sidebar.caption(f"`{Path(draft['path']).name}`")
    else:
        st.sidebar.caption('No draft generated yet.')
    with st.sidebar.expander('Input Examples', expanded=False):
        st.markdown('Chart: `line|Jan:10,Feb:20,Mar:15`')
        st.markdown('Table detail: `tables: Name|Q1|Q2`')
        st.markdown('Hyperlink detail: `hyperlinks: Example Site|https://example.com`')


def _load_recipe_input(uploaded_file, pasted_recipe: str) -> None:
    recipe_text = str(pasted_recipe or '')
    source_name = 'recipe.yaml'

    if uploaded_file is not None:
        source_name = str(getattr(uploaded_file, 'name', 'recipe.yaml') or 'recipe.yaml')
        try:
            recipe_text = uploaded_file.getvalue().decode('utf-8')
        except UnicodeDecodeError:
            st.session_state.recipe_document = None
            st.session_state.recipe_errors = ['Recipe file must be UTF-8 encoded text.']
            st.session_state.recipe_source_name = source_name
            st.session_state.recipe_results = None
            return

    document, errors = load_recipe_text(recipe_text, source_name=source_name)
    st.session_state.recipe_document = document
    st.session_state.recipe_errors = errors
    st.session_state.recipe_source_name = source_name
    st.session_state.recipe_results = None
    st.session_state.file_pack_notes = []
    if document and not errors:
        cards, notes = hydrate_cards_from_document(document)
        st.session_state.file_pack_name = str(document.get('name', '') or 'My File Pack')
        st.session_state.file_pack_cards = cards
        st.session_state.file_pack_notes = notes


def _render_recipe_preview(document: dict[str, object]) -> None:
    items = list(document.get('items', []))
    rows = []
    for index, item in enumerate(items, start=1):
        warning = append_warning_for_item(item)
        rows.append(
            {
                'Item': index,
                'Action': RECIPE_ACTION_LABELS.get(str(item.get('action', '')), str(item.get('action', ''))),
                'Type': str(item.get('file_type', '')).upper(),
                'Target file': str(item.get('file_name', '')),
                'Content preview': summarize_recipe_item(item),
                'Warnings': warning or 'Ready',
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)


def _render_recipe_results(summary: dict[str, object]) -> None:
    metric_col1, metric_col2, metric_col3 = st.columns(3)
    metric_col1.metric('Total items', int(summary.get('total', 0)))
    metric_col2.metric('Successful', int(summary.get('success_count', 0)))
    metric_col3.metric('Failed', int(summary.get('failure_count', 0)))

    for result in summary.get('results', []):
        title = (
            f"Item {result['index']} · {RECIPE_ACTION_LABELS.get(str(result['action']), str(result['action']))}"
            f" · {str(result['file_type']).upper()} · {result['file_name']}"
        )
        with st.container(border=True):
            st.markdown(title)
            if result.get('success'):
                saved_path = str(result.get('saved_path', '')).strip()
                if saved_path:
                    st.success(f'Saved to `{saved_path}`')
                else:
                    st.success('Completed successfully.')
            else:
                st.error(str(result.get('message', 'Unknown error.')))


def _load_file_pack_template(template_key: str) -> None:
    pack_name, cards = build_pack_from_template(template_key)
    st.session_state.file_pack_template = template_key
    st.session_state.file_pack_name = pack_name
    st.session_state.file_pack_cards = cards
    st.session_state.recipe_document = None
    st.session_state.recipe_errors = []
    st.session_state.recipe_results = None
    st.session_state.file_pack_notes = []


def _add_file_pack_card() -> None:
    cards = list(st.session_state.get('file_pack_cards', []))
    cards.append(default_file_pack_card())
    st.session_state.file_pack_cards = cards


def _duplicate_file_pack_card(index: int) -> None:
    cards = list(st.session_state.get('file_pack_cards', []))
    if 0 <= index < len(cards):
        cards.insert(index + 1, dict(cards[index]))
        st.session_state.file_pack_cards = cards


def _remove_file_pack_card(index: int) -> None:
    cards = list(st.session_state.get('file_pack_cards', []))
    if 0 <= index < len(cards):
        del cards[index]
    if not cards:
        cards = [default_file_pack_card()]
    st.session_state.file_pack_cards = cards


def _collect_file_pack_cards_from_state() -> list[dict[str, object]]:
    cards = list(st.session_state.get('file_pack_cards', []))
    collected: list[dict[str, object]] = []
    for index, fallback_card in enumerate(cards):
        prefix = f'pack_card_{index}_'
        collected.append(
            {
                'file_type': st.session_state.get(prefix + 'file_type', fallback_card.get('file_type', 'txt')),
                'file_name': st.session_state.get(prefix + 'file_name', fallback_card.get('file_name', '')),
                'content': st.session_state.get(prefix + 'content', fallback_card.get('content', '')),
                'style': st.session_state.get(prefix + 'style', fallback_card.get('style', '')),
                'font': st.session_state.get(prefix + 'font', fallback_card.get('font', '')),
                'color': st.session_state.get(prefix + 'color', fallback_card.get('color', '')),
                'size': st.session_state.get(prefix + 'size', fallback_card.get('size', 0.0)),
                'alignment': st.session_state.get(prefix + 'alignment', fallback_card.get('alignment', '')),
                'styles_raw': st.session_state.get(prefix + 'styles_raw', fallback_card.get('styles_raw', '')),
                'details_raw': st.session_state.get(prefix + 'details_raw', fallback_card.get('details_raw', '')),
                'chart_type': st.session_state.get(prefix + 'chart_type', fallback_card.get('chart_type', '')),
                'chart_data_raw': st.session_state.get(prefix + 'chart_data_raw', fallback_card.get('chart_data_raw', '')),
            }
        )
    st.session_state.file_pack_cards = collected
    return collected


def _render_file_pack_card(index: int, card: dict[str, object]) -> None:
    prefix = f'pack_card_{index}_'
    safe_file_type = str(card.get('file_type', 'txt'))
    if safe_file_type not in SUPPORTED_FILE_TYPES:
        safe_file_type = 'txt'
    file_type = st.selectbox(
        'File type',
        SUPPORTED_FILE_TYPES,
        index=SUPPORTED_FILE_TYPES.index(safe_file_type),
        key=prefix + 'file_type',
    )
    st.text_input(
        'File name',
        value=str(card.get('file_name', '')),
        key=prefix + 'file_name',
        placeholder='project_brief.docx',
        help='You can leave off the extension and the app will add it.',
    )
    st.text_area(
        'What should this file contain?',
        value=str(card.get('content', '')),
        key=prefix + 'content',
        height=140,
        placeholder='Describe what you want in simple language.',
    )

    if file_type == 'chart':
        chart_cols = st.columns([0.45, 0.55])
        with chart_cols[0]:
            safe_chart_type = str(card.get('chart_type', 'bar'))
            if safe_chart_type not in {'line', 'bar', 'pie', 'scatter'}:
                safe_chart_type = 'bar'
            st.selectbox(
                'Chart type',
                ['line', 'bar', 'pie', 'scatter'],
                index=['line', 'bar', 'pie', 'scatter'].index(safe_chart_type),
                key=prefix + 'chart_type',
            )
        with chart_cols[1]:
            st.text_area(
                'Chart data',
                value=str(card.get('chart_data_raw', '')),
                key=prefix + 'chart_data_raw',
                height=100,
                placeholder='Jan:10,Feb:20,Mar:15',
            )

    with st.expander('Advanced options', expanded=False):
        st.text_input('Style or tone', value=str(card.get('style', '')), key=prefix + 'style')
        format_cols = st.columns(2)
        with format_cols[0]:
            st.text_input('Font', value=str(card.get('font', '')), key=prefix + 'font')
            st.number_input(
                'Size',
                min_value=0.0,
                max_value=120.0,
                step=1.0,
                value=float(card.get('size', 0.0) or 0.0),
                key=prefix + 'size',
            )
        with format_cols[1]:
            st.text_input('Color', value=str(card.get('color', '')), key=prefix + 'color')
            safe_alignment = str(card.get('alignment', ''))
            if safe_alignment not in {'', 'left', 'center', 'right', 'justify'}:
                safe_alignment = ''
            st.selectbox(
                'Alignment',
                ['', 'left', 'center', 'right', 'justify'],
                index=['', 'left', 'center', 'right', 'justify'].index(safe_alignment),
                key=prefix + 'alignment',
            )
        st.text_input('Text styles', value=str(card.get('styles_raw', '')), key=prefix + 'styles_raw')
        st.text_area(
            'Extra details',
            value=str(card.get('details_raw', '')),
            key=prefix + 'details_raw',
            height=120,
            placeholder='headers: Weekly Update\nnotes: keep it short',
        )

    action_cols = st.columns([1, 1, 2.5])
    with action_cols[0]:
        st.button('Duplicate', key=prefix + 'duplicate', on_click=_duplicate_file_pack_card, args=(index,), use_container_width=True)
    with action_cols[1]:
        st.button('Remove', key=prefix + 'remove', on_click=_remove_file_pack_card, args=(index,), use_container_width=True)


def _render_file_pack_tab() -> None:
    st.info('Build a file pack visually, preview what will be created, then create all files in one run.')
    sample_text = recipe_sample_text()
    template_choices = get_starter_template_choices()
    template_map = dict(template_choices)

    st.markdown('### Start a Pack')
    starter_cols = st.columns([1.2, 1.0, 1.0])
    safe_template = str(st.session_state.get('file_pack_template', 'blank'))
    if safe_template not in template_map:
        safe_template = 'blank'
    selected_template = starter_cols[0].selectbox(
        'Starter template',
        options=list(template_map.keys()),
        index=list(template_map.keys()).index(safe_template),
        format_func=lambda key: template_map.get(key, key),
    )
    starter_cols[1].button('Load Template', on_click=_load_file_pack_template, args=(selected_template,), use_container_width=True)
    starter_cols[2].button('Add File', on_click=_add_file_pack_card, use_container_width=True)
    st.text_input('Pack name', key='file_pack_name')

    st.markdown('### Build Your Files')
    current_cards = _collect_file_pack_cards_from_state()
    for index, card in enumerate(current_cards):
        with st.container(border=True):
            st.markdown(f'**File {index + 1}**')
            _render_file_pack_card(index, card)

    current_cards = _collect_file_pack_cards_from_state()
    pack_name = str(st.session_state.get('file_pack_name', 'My File Pack') or 'My File Pack')
    document, pack_errors = build_document_from_cards(pack_name, current_cards)
    st.session_state.recipe_document = document
    st.session_state.recipe_errors = pack_errors

    st.markdown('### Create All Files')
    st.dataframe(preview_file_pack_cards(current_cards), use_container_width=True, hide_index=True)
    for note in st.session_state.get('file_pack_notes', []):
        st.info(note)
    if pack_errors:
        bullets = '\n'.join(f'- {error}' for error in pack_errors)
        st.error(f'Please fix these file pack issues before running:\n{bullets}')

    export_document = document or build_recipe_document(name=pack_name, items=[])
    export_text = dump_recipe_document(export_document)
    action_cols = st.columns([1.0, 1.0, 1.4])
    action_cols[0].download_button(
        'Download Pack',
        data=export_text,
        file_name=f'{Path(pack_name).stem or "file_pack"}.yaml',
        mime='text/yaml',
        use_container_width=True,
    )
    run_disabled = bool(pack_errors) or document is None or not document.get('items')
    if action_cols[1].button('Create All Files', use_container_width=True, disabled=run_disabled):
        progress_bar = st.progress(0.0)
        status = st.empty()

        def _progress_callback(index: int, total: int, result: dict[str, object]) -> None:
            progress_bar.progress(index / total if total else 1.0)
            status.write(f"Created {index}/{total}: {result.get('file_name', '')}")

        summary = run_recipe_batch(list(document.get('items', [])), progress_callback=_progress_callback)
        st.session_state.recipe_results = summary
        progress_bar.progress(1.0)
        status.write('File pack finished.')

    recipe_results = st.session_state.get('recipe_results')
    if recipe_results:
        _render_recipe_results(recipe_results)

    with st.expander('Advanced', expanded=False):
        st.caption('Use this only if you want to import or edit raw YAML/JSON file packs directly.')
        advanced_cols = st.columns([1.4, 1.0])
        with advanced_cols[0]:
            uploaded_recipe = st.file_uploader('Upload file pack', type=['yaml', 'yml', 'json'], key='recipe_file')
            pasted_recipe = st.text_area(
                'Paste YAML or JSON',
                key='recipe_text',
                height=220,
                placeholder='Paste a raw file pack here, then click "Load into Builder".',
            )
            if st.button('Load into Builder', type='primary', use_container_width=True):
                _load_recipe_input(uploaded_recipe, pasted_recipe)
        with advanced_cols[1]:
            st.caption('Sample raw file pack')
            st.code(sample_text, language='yaml')
            st.download_button(
                'Download Sample File Pack',
                data=sample_text,
                file_name='sample_file_pack.yaml',
                mime='text/yaml',
                use_container_width=True,
            )
        recipe_errors = list(st.session_state.get('recipe_errors', []))
        recipe_document = st.session_state.get('recipe_document')
        recipe_source_name = str(st.session_state.get('recipe_source_name', '') or '')
        if recipe_errors:
            bullets = '\n'.join(f'- {error}' for error in recipe_errors)
            st.error(f'Advanced import issues:\n{bullets}')
        elif recipe_document:
            label = str(recipe_document.get('name', '') or recipe_source_name or 'Imported file pack')
            st.success(f'Loaded into builder: {label}')
            _render_recipe_preview(recipe_document)


def build_ui() -> None:
    _init_state()
    st.set_page_config(page_title='File Generator UI', page_icon='📄', layout='wide')
    st.title('File Generator')
    st.caption('Create a draft, review it, request changes, and save only when you are satisfied.')
    _render_sidebar_help()
    if st.sidebar.button('Clear Draft'):
        st.session_state.draft = None
        st.session_state.feedback_text = ''
        st.session_state.feedback_history = []
        st.session_state.recipe_document = None
        st.session_state.recipe_errors = []
        st.session_state.recipe_results = None
        st.session_state.file_pack_name = 'My File Pack'
        st.session_state.file_pack_cards = [default_file_pack_card()]
        st.session_state.file_pack_template = 'blank'
        st.session_state.file_pack_notes = []
        st.rerun()

    create_tab, manage_tab, recipe_tab = st.tabs(['Create / Refine / Save', 'Read / Delete', 'File Packs'])

    with create_tab:
        st.info('Follow the guided flow: choose what to make, add content, tweak the look, preview, then refine.')
        st.markdown('**Workflow:** Create -> Preview -> Request changes -> Save when happy.')
        input_col, config_col = st.columns([1.3, 1.0])

        with input_col:
            st.markdown('### Step 1: Goal & Target')
            create_action_label = st.radio(
                'Draft mode',
                options=list(CREATE_ACTION_LABELS.keys()),
                horizontal=True,
                help='Write = start fresh. Append = continue an existing draft.',
            )
            create_action = CREATE_ACTION_LABELS[create_action_label]
            file_type = st.selectbox(
                'What are we making?',
                SUPPORTED_FILE_TYPES,
                help='Pick the output type first so we can tailor guidance.',
            )
            st.caption(f'Hint: {FILE_TYPE_HINTS.get(file_type, "Describe what you want and any must-have sections.")}')
            output_name = resolve_file_name(
                st.text_input('Save name (no extension needed)', value='output'),
                file_type,
            )
            append_target = ''
            if create_action == 'A':
                suggested_append = ''
                if st.session_state.draft and Path(st.session_state.draft['path']).exists():
                    suggested_append = st.session_state.draft['path']
                append_target = st.text_input(
                    'Append to existing file (optional)',
                    value=suggested_append,
                    help='Leave blank to append to the current draft. Provide a path/name to target another file.',
                )

            st.markdown('### Step 2: Content & Data')
            content = st.text_area(
                'Describe the file or paste source content',
                height=220,
                placeholder='Example: "Create a 1-page project brief with goals, timeline, risks, and next steps."',
            )

            chart_type = ''
            chart_data_raw = ''
            if file_type == 'chart':
                chart_col1, chart_col2 = st.columns([0.48, 0.52])
                with chart_col1:
                    chart_type = st.selectbox('Chart type', ['line', 'bar', 'pie', 'scatter'])
                with chart_col2:
                    chart_data_raw = st.text_area(
                        'Chart data (label:value, comma/line separated)',
                        value='Jan:10,Feb:20,Mar:15',
                        height=120,
                        help='Example: Region A:42, Region B:36. Duplicates are blocked.',
                    )
                st.caption('Tip: You can use percentages or numbers. We validate the format before generating.')

        with config_col:
            st.markdown('### Step 3: Look & Details')
            style = st.text_input(
                'Overall style or tone (optional)',
                value='',
                placeholder='Minimalist, playful, business formal, technical, brand colors, etc.',
            )
            st.caption('Leave blank to use the default style for the chosen file type.')

            with st.expander('Formatting (font, color, size, alignment)', expanded=False):
                font = st.text_input('Font family', value='', placeholder='Inter, Georgia, Consolas...')
                color = st.text_input('Primary color', value='', placeholder='e.g., #0F766E or "indigo 600"')
                size = st.number_input('Base size', min_value=0.0, max_value=120.0, step=1.0, value=0.0)
                alignment = st.selectbox('Alignment', ['', 'left', 'center', 'right', 'justify'])
                styles_raw = st.text_input('Styles (comma-separated)', value='', placeholder='bold, italic, underline')

            with st.expander('Structure & extras', expanded=False):
                details_raw = st.text_area(
                    'Details (one per line: category: value)',
                    height=170,
                    placeholder='tables: Name|Q1|Q2 / A|10|20\nhyperlinks: Example Site|https://example.com\nheaders: Project Falcon - Status',
                )
                st.caption(
                    'Supported keywords: images, tables, charts, hyperlinks, fonts, colors, margins, headers, '
                    'footers, page_numbers, citations, notes, and more. Unknown items are stored as notes.'
                )

        format_options = _collect_format_options(font, color, size, styles_raw, alignment)
        detail_items = parse_details(details_raw)
        errors, warnings = _validate_generation_inputs(
            create_action=create_action,
            file_type=file_type,
            content=content,
            chart_type=chart_type,
            chart_data_raw=chart_data_raw,
            append_target=append_target,
        )

        recipe_export_item = _build_recipe_item_from_form(
            action=create_action,
            file_type=file_type,
            file_name=output_name,
            content=content,
            chart_type=chart_type,
            chart_data_raw=chart_data_raw,
            style=style.strip(),
            format_options=format_options,
            detail_items=detail_items,
        )
        recipe_export_text = dump_recipe_document(
            build_recipe_document(
                name=f'{Path(output_name).stem or "file"} recipe',
                items=[recipe_export_item],
            )
        )

        button_col1, button_col2, button_col3, status_col = st.columns([0.8, 0.8, 1.0, 1.2])
        with button_col1:
            generate_clicked = st.button('Step 4: Generate Preview', type='primary', use_container_width=True)
        with button_col2:
            reset_clicked = st.button('Reset Inputs', use_container_width=True)
        with button_col3:
            st.download_button(
                'Download as File Pack',
                data=recipe_export_text,
                file_name=f'{Path(output_name).stem or "recipe"}.yaml',
                mime='text/yaml',
                use_container_width=True,
                disabled=bool(errors),
                help='Save this setup as a reusable file pack.',
            )
        with status_col:
            _show_step_feedback(errors, warnings, 'Ready: click "Generate Preview" to create a draft.')

        if reset_clicked:
            st.session_state.feedback_text = ''
            st.session_state.feedback_history = []
            st.rerun()

        if generate_clicked:
            if errors:
                st.error('Cannot generate until the required fixes above are addressed.')
            else:
                if create_action == 'A':
                    if append_target.strip():
                        target = append_target.strip()
                        draft_path = resolve_file_name(target, file_type) if '.' not in Path(target).name else target
                        if not Path(draft_path).exists():
                            st.info('Append target not found; creating a new draft instead.')
                            draft_path = _new_draft_path(file_type)
                    elif st.session_state.draft and Path(st.session_state.draft['path']).exists():
                        draft_path = st.session_state.draft['path']
                    else:
                        draft_path = _new_draft_path(file_type)
                        st.info('No existing draft selected. Appending to a new draft file.')
                else:
                    draft_path = _new_draft_path(file_type)

                result = run_action(
                    action=create_action,
                    file_name=draft_path,
                    content=content,
                    chart_type=chart_type if file_type == 'chart' else '',
                    chart_data_raw=chart_data_raw,
                    style=style.strip(),
                    format_options=format_options,
                    detail_items=detail_items,
                )
                if _is_error_text(result):
                    st.error(result)
                else:
                    final_path = Path(result) if Path(str(result)).exists() else Path(draft_path)
                    st.session_state.draft = {
                        'path': str(final_path),
                        'file_type': file_type,
                        'output_name': output_name,
                        'style': style.strip(),
                        'format_options': format_options,
                        'detail_items': detail_items,
                        'chart_type': chart_type,
                    }
                    st.session_state.feedback_history = []
                    st.success('Preview generated. Review and refine below.')

        draft = st.session_state.draft
        if draft and Path(draft['path']).exists():
            st.markdown('---')
            preview_col, review_col = st.columns([1.45, 1.0])

            with preview_col:
                _render_preview(Path(draft['path']), draft['file_type'])

            with review_col:
                st.subheader('Finalize')
                try:
                    draft_path = Path(draft['path'])
                    draft_bytes = draft_path.read_bytes()
                    mime = _guess_mime(draft_path)
                    st.download_button(
                        'Save This File',
                        data=draft_bytes,
                        file_name=draft.get('output_name', draft_path.name),
                        mime=mime,
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f'Unable to prepare download: {exc}')

                st.subheader('Request Changes')
                st.caption('Describe what to change, then apply. You can repeat until satisfied.')
                st.text_area('Change Instructions', key='feedback_text', height=130)
                st.button(
                    'Apply Changes with AI',
                    use_container_width=True,
                    on_click=_handle_apply_changes,
                )
                _render_feedback_notice()

                if st.session_state.feedback_history:
                    st.caption('Applied change requests:')
                    for idx, item in enumerate(st.session_state.feedback_history, start=1):
                        st.write(f'{idx}. {item}')

    with manage_tab:
        st.info('Use this tab for quick reads or deletes without the draft workflow.')
        st.markdown('**Tip:** Add a short focus instruction for reads (e.g., "summarize section 2"). Leave delete target empty to remove the whole file.')
        manage_col1, manage_col2 = st.columns([1.2, 1.0])

        with manage_col1:
            manage_action_label = st.radio(
                'Operation',
                options=list(READ_DELETE_ACTION_LABELS.keys()),
                horizontal=True,
                help='Read shows file content (with optional focus). Delete removes content or the whole file.',
            )
            manage_action = READ_DELETE_ACTION_LABELS[manage_action_label]
            manage_file_type = st.selectbox('File Type', SUPPORTED_FILE_TYPES, key='manage_file_type')
            manage_file_name = resolve_file_name(
                st.text_input('Target file name or path', value='output', key='manage_output_name'),
                manage_file_type,
            )
            st.caption(f'Will target `{manage_file_name}` relative to the current working directory.')
            manage_content = st.text_area(
                'Read focus or delete target (optional)',
                height=180,
                placeholder='Read: "summarize takeaways" or "extract tables". Delete: "table 2" or leave blank to delete file.',
            )
            if manage_action == 'D' and not manage_content.strip():
                st.warning('Delete target is empty. Running delete will remove the entire file.')

        with manage_col2:
            st.subheader('Run')
            st.caption('This executes immediately - no preview draft.')
            if st.button('Run Action', type='primary', use_container_width=True):
                if manage_action == 'R' and not Path(manage_file_name).exists():
                    st.error('File not found. Check the name or provide a full path.')
                elif manage_action == 'D' and not Path(manage_file_name).exists():
                    st.error('File not found. Nothing to delete.')
                else:
                    result = run_action(
                        action=manage_action,
                        file_name=manage_file_name,
                        content=manage_content,
                        chart_type='',
                        chart_data_raw='',
                        style='',
                        format_options={},
                        detail_items=[],
                    )
                    if _is_error_text(result):
                        st.error(result)
                    else:
                        st.success('Completed')
                        st.text_area('Result', value=str(result), height=240, disabled=True)
                        if Path(manage_file_name).exists():
                            st.caption(f'File: `{Path(manage_file_name).resolve()}`')

    with recipe_tab:
        _render_file_pack_tab()


if __name__ == '__main__':
    build_ui()
