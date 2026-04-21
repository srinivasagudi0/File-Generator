from __future__ import annotations

import json
import re
from typing import Any

try:
    import yaml
except ImportError:
    yaml = None

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
SUPPORTED_FILE_TYPES = {
    'txt', 'docx', 'xlsx', 'csv', 'pdf', 'pptx', 'markdown',
    'html', 'code', 'image', 'chart', 'audio', 'video',
}
SUPPORTED_DETAIL_CATEGORIES = {
    'images', 'tables', 'charts', 'graphs', 'hyperlinks',
    'fonts', 'colors', 'sizes', 'styles', 'alignments',
    'margins', 'paddings', 'borders', 'backgrounds', 'layouts',
    'templates', 'themes', 'sections', 'headers', 'footers',
    'page_numbers', 'tables_of_contents', 'indexes', 'bibliographies',
    'citations', 'footnotes', 'notes',
}
ALLOWED_STYLES = {'bold', 'italic', 'underline', 'uppercase', 'lowercase', 'title'}
ALLOWED_ALIGNMENTS = {'left', 'center', 'right', 'justify'}
ACTION_ALIASES = {
    'w': 'W',
    'write': 'W',
    'a': 'A',
    'append': 'A',
}
ACTION_LABELS = {'W': 'Write', 'A': 'Append'}
STARTER_TEMPLATES: dict[str, dict[str, Any]] = {
    'blank': {
        'name': 'My File Pack',
        'items': [
            {
                'action': 'W',
                'file_type': 'txt',
                'file_name': 'notes.txt',
                'content': '',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            }
        ],
    },
    'school_project': {
        'name': 'School Project Pack',
        'items': [
            {
                'action': 'W',
                'file_type': 'docx',
                'file_name': 'project_report.docx',
                'content': 'Create a project report with an introduction, research findings, timeline, and conclusion.',
                'style': 'clear and academic',
                'format_options': {},
                'details': [('headers', 'Project Report')],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'pptx',
                'file_name': 'project_slides.pptx',
                'content': 'Create presentation slides with title, problem, approach, findings, and next steps.',
                'style': 'student presentation',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'txt',
                'file_name': 'sources.txt',
                'content': 'List the sources, article names, and short notes for each source used in the project.',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            },
        ],
    },
    'meeting': {
        'name': 'Meeting Pack',
        'items': [
            {
                'action': 'W',
                'file_type': 'docx',
                'file_name': 'meeting_notes.docx',
                'content': 'Create meeting notes with agenda, discussion points, decisions, and action items.',
                'style': 'professional and concise',
                'format_options': {},
                'details': [('headers', 'Meeting Notes')],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'txt',
                'file_name': 'follow_up_email.txt',
                'content': 'Draft a short follow-up email that summarizes decisions, owners, and deadlines.',
                'style': 'friendly professional',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'csv',
                'file_name': 'action_items.csv',
                'content': 'Owner,Task,Deadline,Status',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            },
        ],
    },
    'club_team': {
        'name': 'Club/Team Pack',
        'items': [
            {
                'action': 'W',
                'file_type': 'docx',
                'file_name': 'weekly_update.docx',
                'content': 'Create a weekly update with wins, challenges, upcoming events, and shout-outs.',
                'style': 'friendly and organized',
                'format_options': {},
                'details': [('headers', 'Weekly Update')],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'chart',
                'file_name': 'attendance.png',
                'content': '',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': 'bar',
                'chart_data_raw': 'Week 1:20,Week 2:18,Week 3:24',
            },
            {
                'action': 'W',
                'file_type': 'markdown',
                'file_name': 'announcement.md',
                'content': 'Write an announcement with this week’s highlights, reminders, and the next meeting time.',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': '',
                'chart_data_raw': '',
            },
        ],
    },
}


def load_recipe_text(raw_text: str, source_name: str = 'recipe') -> tuple[dict[str, Any] | None, list[str]]:
    text = str(raw_text or '')
    if not text.strip():
        return None, ['Recipe is empty. Paste recipe text or upload a recipe file first.']

    parsed, error = _parse_recipe_payload(text, source_name)
    if error:
        return None, [error]
    return normalize_recipe_document(parsed)


def normalize_recipe_document(data: object) -> tuple[dict[str, Any] | None, list[str]]:
    if not isinstance(data, dict):
        return None, ['Recipe must be an object with version, optional name, and items.']

    errors: list[str] = []
    version = data.get('version')
    if version != 1:
        errors.append('Recipe version must be 1.')

    name = str(data.get('name', '') or '').strip()
    raw_items = data.get('items')
    if raw_items is None:
        errors.append('Recipe is missing items.')
        raw_items = []
    elif not isinstance(raw_items, list):
        errors.append('Recipe items must be a list.')
        raw_items = []
    elif not raw_items:
        errors.append('Recipe must include at least one item.')

    items: list[dict[str, Any]] = []
    for index, raw_item in enumerate(raw_items, start=1):
        normalized_item, item_errors = _normalize_recipe_item(raw_item, index)
        errors.extend(item_errors)
        if normalized_item is not None:
            items.append(normalized_item)

    if errors:
        return None, errors

    return {
        'version': 1,
        'name': name,
        'items': items,
    }, []


def dump_recipe_document(document: dict[str, Any]) -> str:
    export_data = _to_export_document(document)
    if yaml is not None:
        return str(yaml.safe_dump(export_data, sort_keys=False, allow_unicode=True))
    return _fallback_yaml_dump(export_data)


def build_recipe_document(
    *,
    name: str = '',
    items: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        'version': 1,
        'name': str(name or '').strip(),
        'items': [dict(item) for item in items],
    }


def recipe_sample_text() -> str:
    sample = build_recipe_document(
        name='Weekly Club Pack',
        items=[
            {
                'action': 'W',
                'file_type': 'docx',
                'file_name': 'meeting_notes.docx',
                'content': 'Create meeting notes with agenda, decisions, and next steps.',
                'style': 'business formal',
                'format_options': {
                    'font': 'Georgia',
                    'color': '#0F766E',
                    'size': 12.0,
                    'styles': ['bold'],
                    'alignment': 'left',
                },
                'details': [('headers', 'Robotics Club')],
                'chart_type': '',
                'chart_data_raw': '',
            },
            {
                'action': 'W',
                'file_type': 'chart',
                'file_name': 'attendance.png',
                'content': '',
                'style': '',
                'format_options': {},
                'details': [],
                'chart_type': 'bar',
                'chart_data_raw': 'Jan:10,Feb:20,Mar:15',
            },
        ],
    )
    return dump_recipe_document(sample)


def summarize_recipe_item(item: dict[str, Any]) -> str:
    if item.get('file_type') == 'chart':
        chart_type = str(item.get('chart_type', '')).strip()
        chart_data = str(item.get('chart_data_raw', '')).strip()
        return f'{chart_type or "chart"} | {chart_data[:80]}'.strip()
    content = str(item.get('content', '')).strip()
    if not content:
        return 'No content provided.'
    if len(content) <= 80:
        return content
    return content[:77].rstrip() + '...'


def default_file_pack_card() -> dict[str, Any]:
    return _card_from_item(dict(STARTER_TEMPLATES['blank']['items'][0]))


def get_starter_template_choices() -> list[tuple[str, str]]:
    return [
        ('blank', 'Blank Pack'),
        ('school_project', 'School Project Pack'),
        ('meeting', 'Meeting Pack'),
        ('club_team', 'Club/Team Pack'),
    ]


def build_pack_from_template(template_key: str) -> tuple[str, list[dict[str, Any]]]:
    template = STARTER_TEMPLATES.get(template_key, STARTER_TEMPLATES['blank'])
    return str(template.get('name', 'My File Pack')), [_card_from_item(dict(item)) for item in template.get('items', [])]


def build_document_from_cards(pack_name: str, cards: list[dict[str, Any]]) -> tuple[dict[str, Any] | None, list[str]]:
    normalized_items: list[dict[str, Any]] = []
    errors: list[str] = []

    for index, card in enumerate(cards, start=1):
        normalized_card, card_errors = normalize_file_pack_card(card, index)
        errors.extend(card_errors)
        if normalized_card is not None:
            normalized_items.append(normalized_card)

    if errors:
        return None, errors
    return build_recipe_document(name=pack_name, items=normalized_items), []


def normalize_file_pack_card(card: dict[str, Any], index: int | None = None) -> tuple[dict[str, Any] | None, list[str]]:
    item_index = index or 1
    if not isinstance(card, dict):
        return None, [f'File {item_index}: invalid file card.']

    raw_file_type = str(card.get('file_type', '') or '').strip().lower()
    file_type = raw_file_type if raw_file_type in SUPPORTED_FILE_TYPES else raw_file_type
    raw_file_name = str(card.get('file_name', '') or '').strip()
    content = str(card.get('content', '') or '')
    style = str(card.get('style', '') or '').strip()
    chart_type = str(card.get('chart_type', '') or '').strip().lower()
    chart_data_raw = str(card.get('chart_data_raw', '') or '').strip()
    detail_items = _details_from_text(str(card.get('details_raw', '') or ''))
    format_options = _card_format_options(card)

    candidate = {
        'action': 'W',
        'file_type': file_type,
        'file_name': resolve_recipe_file_name(raw_file_name, file_type) if raw_file_name and file_type in SUPPORTED_FILE_TYPES else raw_file_name,
        'content': content,
        'style': style,
        'format_options': format_options,
        'details': detail_items,
        'chart_type': chart_type if file_type == 'chart' else '',
        'chart_data_raw': chart_data_raw if file_type == 'chart' else '',
    }

    normalized_item, errors = _normalize_recipe_item(
        {
            'action': 'write',
            'file_type': candidate['file_type'],
            'file_name': candidate['file_name'],
            'content': candidate['content'],
            'style': candidate['style'],
            'format_options': candidate['format_options'],
            'details': [{'category': key, 'value': value} for key, value in candidate['details']],
            'chart_type': candidate['chart_type'],
            'chart_data': candidate['chart_data_raw'],
        },
        item_index,
    )

    friendly_errors = [error.replace('Item', 'File') for error in errors]
    return normalized_item, friendly_errors


def preview_file_pack_cards(cards: list[dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for index, card in enumerate(cards, start=1):
        normalized_item, errors = normalize_file_pack_card(card, index)
        file_type = str(card.get('file_type', '') or '').strip().upper() or 'Unknown'
        file_name = str(card.get('file_name', '') or '').strip() or 'Missing file name'
        summary = summarize_recipe_item(normalized_item or {
            'file_type': str(card.get('file_type', '') or '').strip().lower(),
            'content': str(card.get('content', '') or ''),
            'chart_type': str(card.get('chart_type', '') or ''),
            'chart_data_raw': str(card.get('chart_data_raw', '') or ''),
        })
        warning = 'Ready'
        if errors:
            warning = '; '.join(errors)
        rows.append(
            {
                'File': str(index),
                'Type': file_type,
                'Target file': file_name,
                'What it will make': summary,
                'Status': warning,
            }
        )
    return rows


def hydrate_cards_from_document(document: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    items = list(document.get('items', [])) if isinstance(document, dict) else []
    cards: list[dict[str, Any]] = []
    notes: list[str] = []

    for index, item in enumerate(items, start=1):
        action = str(item.get('action', 'W'))
        if action != 'W':
            notes.append(f'File {index}: append items stay advanced-only and were not loaded into the visual builder.')
            continue
        cards.append(_card_from_item(item))

    if not cards:
        cards.append(default_file_pack_card())
    return cards, notes


def append_warning_for_item(item: dict[str, Any]) -> str:
    if item.get('action') != 'A':
        return ''
    file_name = str(item.get('file_name', '')).strip()
    if file_name and not re.match(r'^[a-zA-Z_]+:', file_name) and not re.match(r'^https?://', file_name):
        from pathlib import Path
        if not Path(file_name).exists():
            return 'Append target does not exist yet; runtime behavior depends on file type.'
    return ''


def _card_from_item(item: dict[str, Any]) -> dict[str, Any]:
    format_options = item.get('format_options', {}) if isinstance(item.get('format_options'), dict) else {}
    styles = format_options.get('styles', [])
    if isinstance(styles, list):
        styles_raw = ', '.join(str(style) for style in styles if str(style).strip())
    else:
        styles_raw = ''
    details = item.get('details', [])
    detail_lines: list[str] = []
    if isinstance(details, list):
        for detail in details:
            if isinstance(detail, (list, tuple)) and len(detail) >= 2:
                detail_lines.append(f'{detail[0]}: {detail[1]}')
            elif isinstance(detail, dict):
                category = str(detail.get('category', '')).strip()
                value = str(detail.get('value', '')).strip()
                if category and value:
                    detail_lines.append(f'{category}: {value}')
    return {
        'file_type': str(item.get('file_type', '') or 'txt'),
        'file_name': str(item.get('file_name', '') or ''),
        'content': str(item.get('content', '') or ''),
        'style': str(item.get('style', '') or ''),
        'font': str(format_options.get('font', '') or ''),
        'color': str(format_options.get('color', '') or ''),
        'size': float(format_options.get('size', 0) or 0),
        'alignment': str(format_options.get('alignment', '') or ''),
        'styles_raw': styles_raw,
        'details_raw': '\n'.join(detail_lines),
        'chart_type': str(item.get('chart_type', '') or ''),
        'chart_data_raw': str(item.get('chart_data_raw', '') or ''),
    }


def _card_format_options(card: dict[str, Any]) -> dict[str, Any]:
    options: dict[str, Any] = {}
    font = str(card.get('font', '') or '').strip()
    color = str(card.get('color', '') or '').strip()
    alignment = str(card.get('alignment', '') or '').strip()
    styles_raw = str(card.get('styles_raw', '') or '')
    size_value = card.get('size', 0)
    if font:
        options['font'] = font
    if color:
        options['color'] = color
    if alignment:
        options['alignment'] = alignment
    try:
        size = float(size_value)
    except (TypeError, ValueError):
        size = 0.0
    if size > 0:
        options['size'] = size
    styles = [piece.strip().lower() for piece in styles_raw.split(',') if piece.strip()]
    if styles:
        options['styles'] = list(dict.fromkeys(styles))
    return options


def _details_from_text(raw_text: str) -> list[tuple[str, str]]:
    detail_items: list[tuple[str, str]] = []
    for line in str(raw_text or '').splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if ':' in stripped:
            category, value = stripped.split(':', 1)
        elif '=' in stripped:
            category, value = stripped.split('=', 1)
        else:
            detail_items.append(('notes', stripped))
            continue
        normalized_category = re.sub(r'[\s\-]+', '_', category.strip().lower())
        normalized_value = value.strip()
        if not normalized_value:
            continue
        if normalized_category not in SUPPORTED_DETAIL_CATEGORIES:
            detail_items.append(('notes', f'{normalized_category}: {normalized_value}'))
            continue
        detail_items.append((normalized_category, normalized_value))
    return detail_items


def _parse_recipe_payload(raw_text: str, source_name: str) -> tuple[object | None, str]:
    stripped = raw_text.lstrip()
    lower_name = str(source_name or '').strip().lower()
    parse_json_first = lower_name.endswith('.json') or stripped.startswith('{')

    if parse_json_first:
        parsed, error = _try_json_then_yaml(raw_text)
    else:
        parsed, error = _try_yaml_then_json(raw_text)
    return parsed, error


def _try_json_then_yaml(raw_text: str) -> tuple[object | None, str]:
    try:
        return json.loads(raw_text), ''
    except json.JSONDecodeError:
        parsed, yaml_error = _try_yaml(raw_text)
        if yaml_error:
            return None, 'Could not parse recipe. Provide valid YAML or JSON.'
        return parsed, ''


def _try_yaml_then_json(raw_text: str) -> tuple[object | None, str]:
    parsed, yaml_error = _try_yaml(raw_text)
    if not yaml_error:
        return parsed, ''
    try:
        return json.loads(raw_text), ''
    except json.JSONDecodeError:
        return None, 'Could not parse recipe. Provide valid YAML or JSON.'


def _try_yaml(raw_text: str) -> tuple[object | None, str]:
    if yaml is None:
        return None, 'YAML recipes require PyYAML. Install `PyYAML` to enable recipe import/export.'
    try:
        return yaml.safe_load(raw_text), ''
    except yaml.YAMLError:
        return None, 'Could not parse YAML recipe.'


def _normalize_recipe_item(raw_item: object, index: int) -> tuple[dict[str, Any] | None, list[str]]:
    if not isinstance(raw_item, dict):
        return None, [f'Item {index}: must be an object.']

    errors: list[str] = []
    action = _normalize_action(raw_item.get('action', 'write'))
    if not action:
        errors.append(f'Item {index}: unsupported action. Use write or append.')
        action = 'W'

    file_type = str(raw_item.get('file_type', '') or '').strip().lower()
    if not file_type:
        errors.append(f'Item {index}: missing file_type.')
    elif file_type not in SUPPORTED_FILE_TYPES:
        errors.append(f'Item {index}: unsupported file_type "{file_type}".')

    raw_file_name = str(raw_item.get('file_name', '') or '').strip()
    if not raw_file_name:
        errors.append(f'Item {index}: missing file_name.')
    resolved_file_name = resolve_recipe_file_name(raw_file_name, file_type) if raw_file_name and file_type in SUPPORTED_FILE_TYPES else raw_file_name

    style = str(raw_item.get('style', '') or '').strip()
    content = '' if raw_item.get('content') is None else str(raw_item.get('content'))

    format_options, format_errors = _normalize_format_options(raw_item.get('format_options'), index)
    errors.extend(format_errors)
    details, detail_errors = _normalize_details(raw_item.get('details'), index)
    errors.extend(detail_errors)

    chart_type = ''
    chart_data_raw = ''
    if file_type == 'chart':
        chart_type = str(raw_item.get('chart_type', '') or '').strip().lower()
        if not chart_type:
            errors.append(f'Item {index}: missing chart_type.')
        elif chart_type not in {'line', 'bar', 'pie', 'scatter'}:
            errors.append(f'Item {index}: chart_type must be line, bar, pie, or scatter.')

        chart_data_raw = _normalize_chart_data(raw_item.get('chart_data'))
        if not chart_data_raw:
            errors.append(f'Item {index}: missing chart_data.')
        else:
            chart_error = _validate_chart_data(chart_data_raw)
            if chart_error:
                errors.append(f'Item {index}: {chart_error}')
    elif not content.strip():
        errors.append(f'Item {index}: missing content.')

    if errors:
        return None, errors

    return {
        'action': action,
        'file_type': file_type,
        'file_name': resolved_file_name,
        'content': content,
        'style': style,
        'format_options': format_options,
        'details': details,
        'chart_type': chart_type,
        'chart_data_raw': chart_data_raw,
    }, []


def _normalize_action(raw_action: object) -> str:
    key = str(raw_action or '').strip().lower()
    return ACTION_ALIASES.get(key, '')


def resolve_recipe_file_name(name: str, file_type: str) -> str:
    cleaned = str(name or '').strip()
    if not cleaned:
        return ''
    if '.' not in cleaned:
        return f'{cleaned}.{DEFAULT_EXT_BY_TYPE[file_type]}'
    return cleaned


def _normalize_format_options(raw_options: object, index: int) -> tuple[dict[str, Any], list[str]]:
    if raw_options is None:
        return {}, []
    if not isinstance(raw_options, dict):
        return {}, [f'Item {index}: format_options must be an object.']

    errors: list[str] = []
    options: dict[str, Any] = {}

    font = str(raw_options.get('font', '') or '').strip()
    if font:
        options['font'] = font

    color = str(raw_options.get('color', '') or '').strip()
    if color:
        options['color'] = color

    raw_size = raw_options.get('size')
    if raw_size not in (None, ''):
        try:
            size = float(raw_size)
        except (TypeError, ValueError):
            errors.append(f'Item {index}: format_options.size must be numeric.')
        else:
            if size <= 0:
                errors.append(f'Item {index}: format_options.size must be greater than 0.')
            else:
                options['size'] = size

    raw_styles = raw_options.get('styles')
    if raw_styles not in (None, ''):
        styles = _normalize_style_list(raw_styles)
        invalid_styles = [style for style in styles if style not in ALLOWED_STYLES]
        if invalid_styles:
            errors.append(f'Item {index}: format_options.styles contains unsupported values.')
        elif styles:
            options['styles'] = styles

    alignment = str(raw_options.get('alignment', '') or '').strip().lower()
    if alignment:
        if alignment not in ALLOWED_ALIGNMENTS:
            errors.append(f'Item {index}: format_options.alignment must be left, center, right, or justify.')
        else:
            options['alignment'] = alignment

    return options, errors


def _normalize_style_list(raw_styles: object) -> list[str]:
    if isinstance(raw_styles, str):
        pieces = [part.strip().lower() for part in raw_styles.split(',') if part.strip()]
    elif isinstance(raw_styles, list):
        pieces = [str(part).strip().lower() for part in raw_styles if str(part).strip()]
    else:
        return []
    return list(dict.fromkeys(pieces))


def _normalize_details(raw_details: object, index: int) -> tuple[list[tuple[str, str]], list[str]]:
    if raw_details is None:
        return [], []

    errors: list[str] = []
    detail_items: list[tuple[str, str]] = []

    if isinstance(raw_details, dict):
        iterable = [{'category': key, 'value': value} for key, value in raw_details.items()]
    elif isinstance(raw_details, list):
        iterable = raw_details
    else:
        return [], [f'Item {index}: details must be a list of category/value objects.']

    for detail_index, entry in enumerate(iterable, start=1):
        category, value, error = _normalize_detail_entry(entry)
        if error:
            errors.append(f'Item {index} detail {detail_index}: {error}')
            continue
        if category and value:
            detail_items.append((category, value))

    return detail_items, errors


def _normalize_detail_entry(entry: object) -> tuple[str, str, str]:
    if isinstance(entry, dict):
        raw_category = entry.get('category', '')
        raw_value = entry.get('value', '')
    elif isinstance(entry, (list, tuple)) and len(entry) >= 2:
        raw_category, raw_value = entry[0], entry[1]
    elif isinstance(entry, str):
        if ':' in entry:
            raw_category, raw_value = entry.split(':', 1)
        elif '=' in entry:
            raw_category, raw_value = entry.split('=', 1)
        else:
            return 'notes', entry.strip(), ''
    else:
        return '', '', 'must include category and value.'

    value = str(raw_value or '').strip()
    if not value:
        return '', '', 'value cannot be empty.'

    category = re.sub(r'[\s\-]+', '_', str(raw_category or '').strip().lower())
    if not category:
        return 'notes', value, ''
    if category not in SUPPORTED_DETAIL_CATEGORIES:
        return 'notes', f'{category}: {value}', ''
    return category, value, ''


def _normalize_chart_data(raw_chart_data: object) -> str:
    if raw_chart_data is None:
        return ''
    if isinstance(raw_chart_data, dict):
        parts = [f'{label}:{value}' for label, value in raw_chart_data.items()]
        return ','.join(parts)
    if isinstance(raw_chart_data, list):
        parts: list[str] = []
        for item in raw_chart_data:
            if isinstance(item, dict) and 'label' in item and 'value' in item:
                parts.append(f"{item['label']}:{item['value']}")
        return ','.join(parts)
    return str(raw_chart_data).strip()


def _validate_chart_data(chart_data_raw: str) -> str:
    entries = [entry.strip() for entry in re.split(r'[,;\n]+', chart_data_raw) if entry.strip()]
    if not entries:
        return 'chart_data must use label:value pairs.'

    seen: set[str] = set()
    for entry in entries:
        if ':' not in entry:
            return 'chart_data must use label:value pairs.'
        label, value = entry.split(':', 1)
        normalized_label = label.strip().casefold()
        if not normalized_label:
            return 'chart_data must use label:value pairs.'
        if normalized_label in seen:
            return f'duplicate chart label "{label.strip()}" is not allowed.'
        try:
            float(value.strip().replace('%', ''))
        except ValueError:
            return f'chart value for "{label.strip()}" must be numeric.'
        seen.add(normalized_label)
    return ''


def _to_export_document(document: dict[str, Any]) -> dict[str, Any]:
    return {
        'version': 1,
        'name': str(document.get('name', '') or '').strip(),
        'items': [_to_export_item(item) for item in document.get('items', [])],
    }


def _to_export_item(item: dict[str, Any]) -> dict[str, Any]:
    export_item: dict[str, Any] = {
        'action': 'write' if item.get('action') == 'W' else 'append',
        'file_type': str(item.get('file_type', '') or '').strip(),
        'file_name': str(item.get('file_name', '') or '').strip(),
    }

    content = str(item.get('content', '') or '')
    if content:
        export_item['content'] = content

    style = str(item.get('style', '') or '').strip()
    if style:
        export_item['style'] = style

    format_options = _export_format_options(item.get('format_options'))
    if format_options:
        export_item['format_options'] = format_options

    details = _export_details(item.get('details'))
    if details:
        export_item['details'] = details

    if str(item.get('file_type', '')).strip() == 'chart':
        export_item['chart_type'] = str(item.get('chart_type', '') or '').strip()
        export_item['chart_data'] = str(item.get('chart_data_raw', '') or '').strip()

    return export_item


def _export_format_options(raw_options: object) -> dict[str, Any]:
    if not isinstance(raw_options, dict):
        return {}
    options: dict[str, Any] = {}
    for key in ('font', 'color', 'alignment'):
        value = str(raw_options.get(key, '') or '').strip()
        if value:
            options[key] = value
    styles = raw_options.get('styles')
    if isinstance(styles, list) and styles:
        options['styles'] = [str(item) for item in styles if str(item).strip()]
    size = raw_options.get('size')
    if isinstance(size, (int, float)):
        options['size'] = int(size) if float(size).is_integer() else float(size)
    return options


def _export_details(raw_details: object) -> list[dict[str, str]]:
    if not isinstance(raw_details, list):
        return []
    exported: list[dict[str, str]] = []
    for item in raw_details:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            category = str(item[0]).strip()
            value = str(item[1]).strip()
        elif isinstance(item, dict):
            category = str(item.get('category', '')).strip()
            value = str(item.get('value', '')).strip()
        else:
            continue
        if category and value:
            exported.append({'category': category, 'value': value})
    return exported


def _fallback_yaml_dump(value: object, indent: int = 0) -> str:
    space = ' ' * indent
    if isinstance(value, dict):
        lines: list[str] = []
        for key, item in value.items():
            if isinstance(item, (dict, list)):
                lines.append(f'{space}{key}:')
                lines.append(_fallback_yaml_dump(item, indent + 2))
            else:
                lines.append(f'{space}{key}: {_yaml_scalar(item)}')
        return '\n'.join(lines)
    if isinstance(value, list):
        lines = []
        for item in value:
            if isinstance(item, (dict, list)):
                nested = _fallback_yaml_dump(item, indent + 2).splitlines()
                first = nested[0] if nested else ''
                lines.append(f'{space}- {first.strip()}')
                for remainder in nested[1:]:
                    lines.append(remainder)
            else:
                lines.append(f'{space}- {_yaml_scalar(item)}')
        return '\n'.join(lines)
    return f'{space}{_yaml_scalar(value)}'


def _yaml_scalar(value: object) -> str:
    if value is None:
        return '""'
    text = str(value)
    if text == '' or re.search(r'[:#\-\n]', text):
        return json.dumps(text)
    return text
