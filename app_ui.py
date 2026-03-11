from __future__ import annotations

import mimetypes
import re
from pathlib import Path

import streamlit as st

from auth_store import (
    allocate_storage_path,
    authenticate_user,
    create_file_record,
    create_user,
    get_file_record,
    is_user_storage_path,
    list_file_records,
    mark_file_deleted,
    sanitize_file_name,
    update_file_record,
)
from file_generator import agent
from intel import process_input as process

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
DELETE_FILE_ALIASES = {'', 'file', 'entire file', 'full file', 'whole file', 'delete file', 'remove file'}
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
    has_append_target: bool,
) -> tuple[list[str], list[str]]:
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

    if create_action == 'A' and not has_append_target:
        warnings.append('No existing private file selected. Append will create a new file in your workspace.')

    return errors, warnings


def _show_step_feedback(errors: list[str], warnings: list[str], ready_text: str) -> None:
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
            return error
        if not chart_data:
            return 'Chart data cannot be empty.'
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


def _init_state() -> None:
    if 'auth_user' not in st.session_state:
        st.session_state.auth_user = None
    if 'draft' not in st.session_state:
        st.session_state.draft = None
    if 'feedback_text' not in st.session_state:
        st.session_state.feedback_text = ''
    if 'feedback_history' not in st.session_state:
        st.session_state.feedback_history = []


def _clear_draft_state() -> None:
    st.session_state.draft = None
    st.session_state.feedback_text = ''
    st.session_state.feedback_history = []


def _clear_user_session() -> None:
    st.session_state.auth_user = None
    _clear_draft_state()


def _guess_mime(path: Path) -> str:
    guessed, _ = mimetypes.guess_type(str(path))
    if guessed:
        return guessed
    return 'application/octet-stream'


def _render_preview(path: Path, file_type: str, label: str) -> None:
    st.subheader('Preview')
    st.caption(f'File: `{label}`')
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


def _is_full_delete_request(value: str) -> bool:
    return str(value or '').strip().lower() in DELETE_FILE_ALIASES


def _record_exists(record: dict[str, object]) -> bool:
    return Path(str(record['storage_path'])).exists()


def _active_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    return [record for record in records if record['status'] == 'active']


def _existing_active_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    return [record for record in _active_records(records) if _record_exists(record)]


def _format_record_label(record: dict[str, object]) -> str:
    status = str(record['status'])
    updated = str(record['updated_at']).replace('T', ' ')
    if status == 'active' and not _record_exists(record):
        status = 'missing'
    return f"{record['display_name']} [{record['file_type']}] - {status} - {updated}"


def _owned_draft(user_id: int) -> dict[str, object] | None:
    draft = st.session_state.draft
    if not draft:
        return None
    if int(draft.get('user_id', -1)) != int(user_id):
        _clear_draft_state()
        return None
    record_id = draft.get('record_id')
    if not isinstance(record_id, int):
        _clear_draft_state()
        return None
    record = get_file_record(user_id, record_id)
    if record is None:
        _clear_draft_state()
        return None
    draft['path'] = str(record['storage_path'])
    draft['output_name'] = str(record['display_name'])
    draft['file_type'] = str(record['file_type'])
    return draft


def _set_draft_from_record(
    user_id: int,
    record: dict[str, object],
    *,
    style: str = '',
    format_options: dict[str, object] | None = None,
    detail_items: list[tuple[str, str]] | None = None,
    chart_type: str = '',
) -> None:
    st.session_state.draft = {
        'user_id': int(user_id),
        'record_id': int(record['id']),
        'path': str(record['storage_path']),
        'file_type': str(record['file_type']),
        'output_name': str(record['display_name']),
        'style': style,
        'format_options': format_options or {},
        'detail_items': detail_items or [],
        'chart_type': chart_type,
    }


def _resolve_owned_result_path(user_id: int, requested_path: Path, result: str) -> tuple[Path | None, str]:
    candidate = Path(str(result).strip()) if isinstance(result, str) and str(result).strip() else requested_path
    final_path = candidate if candidate.exists() else requested_path
    final_path = final_path.resolve()
    if not is_user_storage_path(user_id, final_path):
        return None, 'Generated file escaped the signed-in user workspace.'
    return final_path, ''


def _apply_feedback_to_draft(
    user_id: int,
    feedback: str,
    style: str,
    format_options: dict[str, object],
    detail_items: list[tuple[str, str]],
) -> tuple[bool, str]:
    draft = _owned_draft(user_id)
    if not draft:
        return False, 'No draft available yet.'

    record = get_file_record(user_id, int(draft['record_id']))
    if record is None:
        return False, 'Draft record not found for the signed-in user.'

    path = Path(str(record['storage_path']))
    if not path.exists():
        return False, 'Draft file not found. Generate a preview again.'

    file_type = str(record['file_type'])
    ext = path.suffix.lower().lstrip('.')

    if file_type == 'image' or ext in IMAGE_EXTENSIONS:
        result = agent(feedback, str(path), 'A')
        if _is_error_text(result):
            return False, str(result)
        final_path, path_error = _resolve_owned_result_path(user_id, path, str(result))
        if path_error or final_path is None or not final_path.exists():
            return False, path_error or 'Image update did not produce a file in your workspace.'
        update_file_record(user_id, int(record['id']), storage_path=final_path)
        draft['path'] = str(final_path)
        return True, str(result)

    if file_type == 'chart':
        chart_type, chart_data_raw, err = _parse_chart_feedback(feedback, str(draft.get('chart_type', '')))
        if err:
            return False, err
        chart_data, parse_error = parse_chart_data(chart_data_raw)
        if parse_error:
            return False, parse_error
        result = agent('', str(path), 'W', chart_type=chart_type, chart_data=chart_data)
        if _is_error_text(result):
            return False, str(result)
        final_path, path_error = _resolve_owned_result_path(user_id, path, str(result))
        if path_error or final_path is None or not final_path.exists():
            return False, path_error or 'Chart update did not produce a file in your workspace.'
        updated = update_file_record(user_id, int(record['id']), storage_path=final_path)
        if updated is None:
            return False, 'Unable to update draft history.'
        draft['chart_type'] = chart_type
        draft['path'] = str(final_path)
        return True, str(result)

    if file_type in ('audio', 'video'):
        return False, 'Feedback-based editing is not supported for audio/video yet. Generate a new file instead.'

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
    if _is_error_text(result):
        return False, str(result)
    final_path, path_error = _resolve_owned_result_path(user_id, path, str(result))
    if path_error or final_path is None or not final_path.exists():
        return False, path_error or 'Updated content did not produce a file in your workspace.'
    updated = update_file_record(user_id, int(record['id']), storage_path=final_path)
    if updated is None:
        return False, 'Unable to update draft history.'
    draft['path'] = str(final_path)
    return True, str(result)


def _render_auth_screen() -> None:
    st.title('File Generator')
    st.caption(
        'Sign in to use a private workspace. Generated files are stored per account, and only that account can '
        'preview or download them.'
    )

    login_tab, register_tab = st.tabs(['Sign In', 'Create Account'])

    with login_tab:
        with st.form('login_form'):
            username = st.text_input('Username', placeholder='yourname')
            password = st.text_input('Password', type='password')
            submitted = st.form_submit_button('Sign In', use_container_width=True)
        if submitted:
            user, error = authenticate_user(username, password)
            if error:
                st.error(error)
            else:
                st.session_state.auth_user = user
                _clear_draft_state()
                st.rerun()

    with register_tab:
        with st.form('register_form'):
            username = st.text_input('New username', placeholder='yourname')
            password = st.text_input('New password', type='password')
            confirm = st.text_input('Confirm password', type='password')
            submitted = st.form_submit_button('Create Account', use_container_width=True)
        if submitted:
            if password != confirm:
                st.error('Passwords do not match.')
            else:
                user, error = create_user(username, password)
                if error:
                    st.error(error)
                else:
                    st.session_state.auth_user = user
                    _clear_draft_state()
                    st.success('Account created.')
                    st.rerun()


def _render_sidebar(user: dict[str, object], records: list[dict[str, object]]) -> None:
    draft = _owned_draft(int(user['id']))
    active_count = len(_existing_active_records(records))
    st.sidebar.header('Account')
    st.sidebar.caption(f"Signed in as `{user['username']}`")
    st.sidebar.caption(f'{active_count} file(s) in your history')

    if st.sidebar.button('Log Out', use_container_width=True):
        _clear_user_session()
        st.rerun()

    if st.sidebar.button('Clear Current Draft', use_container_width=True):
        _clear_draft_state()
        st.rerun()

    st.sidebar.header('Workspace')
    st.sidebar.markdown('1. Generate into your private storage')
    st.sidebar.markdown('2. Refine the current draft')
    st.sidebar.markdown('3. Reopen any file from My Files')
    st.sidebar.markdown('4. Read or delete only owned files')
    if draft and Path(str(draft['path'])).exists():
        st.sidebar.success('Draft ready')
        st.sidebar.caption(f"`{draft['output_name']}`")
    else:
        st.sidebar.caption('No active draft selected.')
    with st.sidebar.expander('Input Examples', expanded=False):
        st.markdown('Chart: `line|Jan:10,Feb:20,Mar:15`')
        st.markdown('Table detail: `tables: Name|Q1|Q2`')
        st.markdown('Hyperlink detail: `hyperlinks: Example Site|https://example.com`')


def _history_rows(records: list[dict[str, object]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for record in records:
        status = str(record['status'])
        if status == 'active' and not _record_exists(record):
            status = 'missing'
        rows.append(
            {
                'File': str(record['display_name']),
                'Type': str(record['file_type']),
                'Status': status,
                'Updated (UTC)': str(record['updated_at']).replace('T', ' '),
            }
        )
    return rows


def build_ui() -> None:
    st.set_page_config(page_title='File Generator UI', page_icon='📄', layout='wide')
    _init_state()

    user = st.session_state.auth_user
    if not user:
        _render_auth_screen()
        return

    user_id = int(user['id'])
    records = list_file_records(user_id, include_deleted=True)
    active_records = _existing_active_records(records)
    draft = _owned_draft(user_id)

    st.title('File Generator')
    st.caption('Authenticated mode: every generated file is saved into the signed-in user workspace and tracked in history.')
    _render_sidebar(user, records)

    create_tab, history_tab, manage_tab = st.tabs(['Create / Refine', 'My Files', 'Read / Delete'])

    with create_tab:
        st.info('Create or refine files inside your own private workspace. File paths are no longer shared across users.')
        st.markdown('**Workflow:** Generate -> Review -> Request changes -> Download from your own history.')
        input_col, config_col = st.columns([1.3, 1.0])

        with input_col:
            st.markdown('### Step 1: Goal & Target')
            create_action_label = st.radio(
                'Draft mode',
                options=list(CREATE_ACTION_LABELS.keys()),
                horizontal=True,
                help='Write = create a new private file. Append = continue an owned file.',
            )
            create_action = CREATE_ACTION_LABELS[create_action_label]
            file_type = st.selectbox(
                'What are we making?',
                SUPPORTED_FILE_TYPES,
                help='Pick the output type first so we can tailor guidance.',
            )
            st.caption(f'Hint: {FILE_TYPE_HINTS.get(file_type, "Describe what you want and any must-have sections.")}')

            default_name = draft['output_name'] if draft and draft['file_type'] == file_type else 'output'
            requested_name = resolve_file_name(
                st.text_input('Save name', value=default_name),
                file_type,
            )
            output_name = sanitize_file_name(requested_name, f'output.{DEFAULT_EXT_BY_TYPE[file_type]}')

            append_options = [
                record for record in active_records
                if str(record['file_type']) == file_type
            ]
            current_draft_available = bool(draft and draft['file_type'] == file_type and Path(str(draft['path'])).exists())
            append_choices: list[str] = ['']
            if current_draft_available:
                append_choices.append('__current__')
            append_choices.extend(str(record['id']) for record in append_options)
            append_target = ''
            if create_action == 'A':
                append_target = st.selectbox(
                    'Append to',
                    options=append_choices,
                    format_func=lambda value: (
                        'Create a new private file'
                        if value == ''
                        else 'Current draft'
                        if value == '__current__'
                        else _format_record_label(next(record for record in append_options if str(record['id']) == value))
                    ),
                    help='Only files in your account history are available here.',
                )
                if append_target not in {'', '__current__'}:
                    st.caption('Append keeps the selected file name and ownership.')

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
                font = st.text_input('Font family', value='', placeholder='Georgia, Consolas...')
                color = st.text_input('Primary color', value='', placeholder='e.g., #0F766E or "teal 700"')
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
            has_append_target=append_target in {'__current__'} or bool(append_target),
        )

        button_col1, button_col2, status_col = st.columns([0.9, 0.9, 1.2])
        with button_col1:
            generate_clicked = st.button('Step 4: Generate Preview', type='primary', use_container_width=True)
        with button_col2:
            reset_clicked = st.button('Reset Inputs', use_container_width=True)
        with status_col:
            _show_step_feedback(errors, warnings, 'Ready: generate a preview into your private history.')

        if reset_clicked:
            st.session_state.feedback_text = ''
            st.session_state.feedback_history = []
            st.rerun()

        if generate_clicked:
            if errors:
                st.error('Cannot generate until the required fixes above are addressed.')
            else:
                existing_record: dict[str, object] | None = None
                if create_action == 'A':
                    if append_target == '__current__' and draft:
                        existing_record = get_file_record(user_id, int(draft['record_id']))
                    elif append_target:
                        existing_record = get_file_record(user_id, int(append_target))

                target_path = (
                    Path(str(existing_record['storage_path']))
                    if existing_record is not None
                    else allocate_storage_path(user_id, output_name)
                )
                result = run_action(
                    action=create_action,
                    file_name=str(target_path),
                    content=content,
                    chart_type=chart_type if file_type == 'chart' else '',
                    chart_data_raw=chart_data_raw,
                    style=style.strip(),
                    format_options=format_options,
                    detail_items=detail_items,
                )
                if _is_error_text(result):
                    if existing_record is None and target_path.exists():
                        try:
                            target_path.unlink()
                        except Exception:
                            pass
                    st.error(result)
                else:
                    final_path, path_error = _resolve_owned_result_path(user_id, target_path, result)
                    if path_error or final_path is None or not final_path.exists():
                        if existing_record is None and target_path.exists():
                            try:
                                target_path.unlink()
                            except Exception:
                                pass
                        st.error(path_error or 'Generation finished without creating a file in your workspace.')
                    else:
                        if existing_record is None:
                            record = create_file_record(
                                user_id,
                                output_name,
                                file_type,
                                final_path,
                            )
                        else:
                            record = update_file_record(
                                user_id,
                                int(existing_record['id']),
                                storage_path=final_path,
                            )
                            if record is None:
                                st.error('Unable to update file history.')
                                record = None

                        if record is not None:
                            _set_draft_from_record(
                                user_id,
                                record,
                                style=style.strip(),
                                format_options=format_options,
                                detail_items=detail_items,
                                chart_type=chart_type,
                            )
                            st.session_state.feedback_history = []
                            st.success('Preview generated and stored in your private history.')
                            st.rerun()

        draft = _owned_draft(user_id)
        if draft:
            draft_record = get_file_record(user_id, int(draft['record_id']))
        else:
            draft_record = None
        if draft_record and Path(str(draft_record['storage_path'])).exists():
            st.markdown('---')
            preview_col, review_col = st.columns([1.45, 1.0])

            with preview_col:
                _render_preview(
                    Path(str(draft_record['storage_path'])),
                    str(draft_record['file_type']),
                    str(draft_record['display_name']),
                )

            with review_col:
                st.subheader('Download')
                st.caption('This download is restricted to the currently signed-in user file.')
                try:
                    draft_path = Path(str(draft_record['storage_path']))
                    draft_bytes = draft_path.read_bytes()
                    mime = _guess_mime(draft_path)
                    st.download_button(
                        'Download Current File',
                        data=draft_bytes,
                        file_name=str(draft_record['display_name']),
                        mime=mime,
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f'Unable to prepare download: {exc}')

                st.subheader('Request Changes')
                st.caption('Describe what to change, then apply. Updates stay on your copy only.')
                feedback = st.text_area('Change Instructions', key='feedback_text', height=130)
                if st.button('Apply Changes with AI', use_container_width=True):
                    if not feedback.strip():
                        st.warning('Enter change instructions first.')
                    else:
                        success, message = _apply_feedback_to_draft(
                            user_id=user_id,
                            feedback=feedback.strip(),
                            style=str(draft.get('style', '')),
                            format_options=dict(draft.get('format_options', {})),
                            detail_items=list(draft.get('detail_items', [])),
                        )
                        st.session_state.feedback_history.append(feedback.strip())
                        if success:
                            st.success('Draft updated. Review the new preview.')
                            st.session_state.feedback_text = ''
                            st.rerun()
                        else:
                            st.error(message)

                if st.session_state.feedback_history:
                    st.caption('Applied change requests:')
                    for idx, item in enumerate(st.session_state.feedback_history, start=1):
                        st.write(f'{idx}. {item}')

    with history_tab:
        st.info('History is account-scoped. Only files recorded for the signed-in user can be previewed or downloaded here.')
        if not records:
            st.caption('No files in your history yet.')
        else:
            st.dataframe(_history_rows(records), hide_index=True, use_container_width=True)

        previewable_records = _existing_active_records(records)
        if previewable_records:
            selected_history_id = st.selectbox(
                'Select a file from your history',
                options=[''] + [str(record['id']) for record in previewable_records],
                format_func=lambda value: (
                    'Choose a file'
                    if value == ''
                    else _format_record_label(next(record for record in previewable_records if str(record['id']) == value))
                ),
            )
            if selected_history_id:
                selected_record = get_file_record(user_id, int(selected_history_id))
                if selected_record and Path(str(selected_record['storage_path'])).exists():
                    preview_col, info_col = st.columns([1.45, 1.0])
                    with preview_col:
                        _render_preview(
                            Path(str(selected_record['storage_path'])),
                            str(selected_record['file_type']),
                            str(selected_record['display_name']),
                        )
                    with info_col:
                        st.subheader('Owned File')
                        st.caption(f"Created: {selected_record['created_at']}")
                        st.caption(f"Updated: {selected_record['updated_at']}")
                        download_path = Path(str(selected_record['storage_path']))
                        try:
                            st.download_button(
                                'Download From History',
                                data=download_path.read_bytes(),
                                file_name=str(selected_record['display_name']),
                                mime=_guess_mime(download_path),
                                use_container_width=True,
                            )
                        except Exception as exc:
                            st.error(f'Unable to prepare download: {exc}')
                        if st.button('Open As Current Draft', use_container_width=True):
                            _set_draft_from_record(user_id, selected_record)
                            st.success('Loaded into the editor.')
                            st.rerun()
        else:
            st.caption('No active files are currently available for preview.')

    with manage_tab:
        st.info('Read and delete are now ownership-checked. You can only target files from your own history.')
        manageable_records = _existing_active_records(records)
        if not manageable_records:
            st.caption('No active owned files available.')
        else:
            manage_col1, manage_col2 = st.columns([1.2, 1.0])
            current_draft_id = str(draft['record_id']) if draft else ''

            with manage_col1:
                manage_action_label = st.radio(
                    'Operation',
                    options=list(READ_DELETE_ACTION_LABELS.keys()),
                    horizontal=True,
                    help='Read shows file content or a summary. Delete removes content or the full owned file.',
                )
                manage_action = READ_DELETE_ACTION_LABELS[manage_action_label]
                manage_file_id = st.selectbox(
                    'Owned file',
                    options=[str(record['id']) for record in manageable_records],
                    index=next(
                        (idx for idx, record in enumerate(manageable_records) if str(record['id']) == current_draft_id),
                        0,
                    ),
                    format_func=lambda value: _format_record_label(
                        next(record for record in manageable_records if str(record['id']) == value)
                    ),
                )
                manage_content = st.text_area(
                    'Read focus or delete target (optional)',
                    height=180,
                    placeholder='Read: "summarize takeaways" or "extract tables". Delete: "table 2" or leave blank to delete file.',
                )
                if manage_action == 'D' and not manage_content.strip():
                    st.warning('Delete target is empty. Running delete will remove the entire owned file.')

            with manage_col2:
                st.subheader('Run')
                st.caption('This executes immediately, but still only on your own file.')
                if st.button('Run Action', type='primary', use_container_width=True):
                    record = get_file_record(user_id, int(manage_file_id))
                    if record is None:
                        st.error('Selected file is no longer available for this account.')
                    else:
                        path = Path(str(record['storage_path']))
                        if not path.exists():
                            st.error('File not found on disk for this history record.')
                        else:
                            result = run_action(
                                action=manage_action,
                                file_name=str(path),
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
                                if manage_action == 'D' and _is_full_delete_request(manage_content):
                                    mark_file_deleted(user_id, int(record['id']))
                                    if draft and int(draft['record_id']) == int(record['id']):
                                        _clear_draft_state()
                                    st.success('Completed and removed the file from your active workspace.')
                                else:
                                    update_file_record(user_id, int(record['id']), storage_path=path)
                                    st.success('Completed')
                                st.text_area('Result', value=str(result), height=240, disabled=True)


if __name__ == '__main__':
    build_ui()
