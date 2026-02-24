

from __future__ import annotations

import re
import sys
import time
from pathlib import Path

from file_generator import agent
from intel import process_input as process

IMAGE_EXTENSIONS = ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp')
PROCESS_SUPPORTED_EXTENSIONS = {'txt', 'docx', 'xlsx', 'xls'}
TEXT_STYLE_EXTENSIONS = {'txt', 'docx'}
CLOUD_STORAGE_PREFIXES = {
    'google_drive', 'googledrive', 'gdrive', 'drive', 'dropbox', 'onedrive', 'icloud',
    'box', 's3', 'amazon_s3', 'pcloud', 'mega', 'sync', 'sync_com', 'nextcloud', 'owncloud', 'tresorit'
}
SUPPORTED_FONTS = (
    'arial', 'times new roman', 'calibri', 'verdana', 'georgia',
    'garamond', 'comic sans ms', 'courier new', 'tahoma', 'helvetica'
)
SUPPORTED_DETAIL_CATEGORIES = (
    'images', 'tables', 'charts', 'graphs', 'hyperlinks',
    'fonts', 'colors', 'sizes', 'styles', 'alignments',
    'margins', 'paddings', 'borders', 'backgrounds', 'layouts',
    'templates', 'themes', 'sections', 'headers', 'footers',
    'page_numbers', 'tables_of_contents', 'indexes', 'bibliographies',
    'citations', 'footnotes', 'notes'
)
DETAIL_CATEGORY_ALIASES = {
    'image': 'images',
    'images': 'images',
    'picture': 'images',
    'pictures': 'images',
    'photo': 'images',
    'photos': 'images',
    'table': 'tables',
    'tables': 'tables',
    'chart': 'charts',
    'charts': 'charts',
    'graph': 'graphs',
    'graphs': 'graphs',
    'plot': 'graphs',
    'plots': 'graphs',
    'hyperlink': 'hyperlinks',
    'hyperlinks': 'hyperlinks',
    'link': 'hyperlinks',
    'links': 'hyperlinks',
    'url': 'hyperlinks',
    'urls': 'hyperlinks',
    'font': 'fonts',
    'fonts': 'fonts',
    'color': 'colors',
    'colors': 'colors',
    'size': 'sizes',
    'sizes': 'sizes',
    'style': 'styles',
    'styles': 'styles',
    'alignment': 'alignments',
    'alignments': 'alignments',
    'margin': 'margins',
    'margins': 'margins',
    'padding': 'paddings',
    'paddings': 'paddings',
    'border': 'borders',
    'borders': 'borders',
    'background': 'backgrounds',
    'backgrounds': 'backgrounds',
    'layout': 'layouts',
    'layouts': 'layouts',
    'template': 'templates',
    'templates': 'templates',
    'theme': 'themes',
    'themes': 'themes',
    'section': 'sections',
    'sections': 'sections',
    'header': 'headers',
    'headers': 'headers',
    'footer': 'footers',
    'footers': 'footers',
    'page_number': 'page_numbers',
    'page_numbers': 'page_numbers',
    'page number': 'page_numbers',
    'page numbers': 'page_numbers',
    'toc': 'tables_of_contents',
    'table_of_contents': 'tables_of_contents',
    'table of contents': 'tables_of_contents',
    'tables_of_contents': 'tables_of_contents',
    'tables of contents': 'tables_of_contents',
    'index': 'indexes',
    'indexes': 'indexes',
    'bibliography': 'bibliographies',
    'bibliographies': 'bibliographies',
    'citation': 'citations',
    'citations': 'citations',
    'footnote': 'footnotes',
    'footnotes': 'footnotes',
    'note': 'notes',
    'notes': 'notes',
}
MAX_DETAIL_ITEMS = 40
MAX_DETAIL_VALUE_LENGTH = 1200
MAX_DETAIL_TOTAL_LENGTH = 12000


def main():
    welcome()
    while True:
        action = action_type()
        atype = file_type()
        chart_type = None
        chart_data = None
        detail_items: list[tuple[str, str]] = []

        if atype == 'chart' and action in ('W', 'A'):
            content = ''
            chart_type, chart_data = ask_chart_inputs()
        else:
            content, detail_items = ask_content(action)

        name = file_name(action, atype)
        ext = Path(name).suffix.lower().lstrip('.')
        style = ''
        format_options: dict[str, object] = {}
        if action in ('W', 'A') and ext in TEXT_STYLE_EXTENSIONS:
            style, format_options = ask_formatting_options(ext)
            format_options = _merge_format_options_from_details(format_options, detail_items)
            if ext == 'txt' and not style and format_options.get('font'):
                style = str(format_options.get('font'))

        print('Got it — processing your request now...')

        try:
            is_image = ext in IMAGE_EXTENSIONS

            if action == 'D':
                result = agent(content, name, action, '')
                if isinstance(result, str) and _is_failure(result):
                    print(result)
                    continue
                print(result)
            elif action == 'R':
                if atype == 'chart':
                    summary = process(content, action, name)
                    if not isinstance(summary, str) or not summary.strip() or _is_failure(summary):
                        summary = _chart_read_fallback(name)
                    print('Here is what I found in your file:')
                    print(summary)
                elif is_image:
                    summary = process(content, action, name)
                else:
                    read_result = agent('', name, action, '')
                    if isinstance(read_result, str) and _is_failure(read_result):
                        print(read_result)
                        continue
                    if _supports_ai_processing(ext):
                        summary = process(content, action, name)
                    else:
                        summary = read_result

                if atype != 'chart' and isinstance(summary, str) and _is_failure(summary):
                    print(summary)
                    continue
                if atype != 'chart':
                    print('Here is what I found in your file:')
                    print(summary)
            elif atype == 'chart':
                result = agent(content, name, action, '', chart_type=chart_type, chart_data=chart_data)
                if isinstance(result, str) and _is_failure(result):
                    print(result)
                    continue
                print('Generating your file...')
                time.sleep(1)
                resolved_path = str(Path(result).resolve()) if isinstance(result, str) and result.strip() else str(Path(name).resolve())
                print(f'All set! Saved to {resolved_path}.')
            else:
                if is_image:
                    payload = content
                elif _supports_ai_processing(ext):
                    payload = process(content, action, name)
                else:
                    payload = content
                if isinstance(payload, str) and _is_failure(payload):
                    print(payload)
                    continue

                result = agent(payload, name, action, style, format_options=format_options, details=detail_items)
                if isinstance(result, str) and _is_failure(result):
                    print(result)
                    continue

                print('Generating your file...')
                time.sleep(1)
                if isinstance(result, str) and result.strip():
                    resolved_path = str(Path(result).resolve()) if Path(result).suffix else result
                    print(f'All set! Saved to {resolved_path}.')
                else:
                    print(f'All set! Saved to {Path.cwd()}.')
        except Exception:
            print('Sorry, something went wrong while processing your request. Please try again or type "help".')
            continue

        if not ask_yes_no('Would you like to generate another file? (yes/no)'):
            print('Thanks for using File Generator — talk soon!')
            break
        # Loop back to start of while loop to generate another file
        print('---------------------------------------')
        continue

def welcome():
    print('Hey there! I\'m your File Generator.')
    print('Type "help" anytime to see what I can do.')
    print("Let's create, read, or tidy up some files together.")
    print('---------------------------------------')
    print('Type "quit" to exit.')

def ask_content(action: str) -> tuple[str, list[tuple[str, str]]]:
    if action == 'R':
        print('What should I focus on when reading the file? (leave blank for a quick summary)')
        return prompt_user('You: '), []

    if action == 'D':
        print('What would you like me to delete from this file?')
        print('Leave blank to delete the entire file.')
        print('For DOCX tables: use "table", "table:2", or "table:contains:keyword".')
        print('For text-only delete: type the text to remove. Use "text:your text" to be explicit.')
        return prompt_user('You: '), []

    if action in ('W', 'A'):
        while True:
            print('What would you like me to generate for you today?')
            print()
            user_input = prompt_user('You: ')
            if not user_input.strip():
                print('No content provided. Try again with a short description.')
                continue

            while True:
                print('Add more details to your request? (yes/no)')
                confirmation = prompt_user('You:').strip().lower()
                if confirmation in ('yes', 'y'):
                    detail_items = ask_generation_details()
                    return user_input, detail_items
                if confirmation in ('no', 'n', ''):
                    return user_input, []
                print('Please answer with yes or no.')

    return '', []

def ask_generation_details() -> list[tuple[str, str]]:
    print('Add details using "category: value" or "category=value".')
    print('Use one detail per line. Press Enter on an empty line when you are done.')
    print('Examples:')
    print('- tables: Name|Q1|Q2 / A|12|16 / B|10|14')
    print('- charts: line|Jan:10,Feb:12,Mar:9')
    print('- images: path:./assets/logo.png or prompt:a blue abstract banner')
    print('- hyperlinks: Hack Club|https://hackclub.com')
    print(f'Supported categories: {", ".join(SUPPORTED_DETAIL_CATEGORIES)}')
    items: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    total_chars = 0

    while len(items) < MAX_DETAIL_ITEMS and total_chars < MAX_DETAIL_TOTAL_LENGTH:
        raw = prompt_user('Detail: ')
        if not raw.strip():
            break
        chunks = _split_detail_chunks(raw)
        if not chunks:
            print('No valid detail found. Try "category: value".')
            continue
        for chunk in chunks:
            category, value, error = _parse_detail_entry(chunk)
            if error:
                print(error)
                continue
            if not category or not value:
                continue
            key = (category, value.casefold())
            if key in seen:
                continue
            seen.add(key)
            items.append((category, value))
            total_chars += len(category) + len(value)
            if len(items) >= MAX_DETAIL_ITEMS or total_chars >= MAX_DETAIL_TOTAL_LENGTH:
                break

    if len(items) >= MAX_DETAIL_ITEMS:
        print(f'Detail limit reached ({MAX_DETAIL_ITEMS} items).')
    if total_chars >= MAX_DETAIL_TOTAL_LENGTH:
        print(f'Detail character limit reached ({MAX_DETAIL_TOTAL_LENGTH}).')
    return items

def _split_detail_chunks(raw: str) -> list[str]:
    chunks = [part.strip() for part in raw.split(';')]
    return [chunk for chunk in chunks if chunk]

def _parse_detail_entry(entry: str) -> tuple[str, str, str]:
    separator_index = -1
    for candidate in (':', '='):
        idx = entry.find(candidate)
        if idx > 0:
            separator_index = idx
            break
    if separator_index < 0:
        return '', '', f'Invalid detail "{entry}". Use "category: value".'

    raw_category = entry[:separator_index].strip()
    raw_value = entry[separator_index + 1:].strip()
    category = _normalize_detail_category(raw_category)
    value = _normalize_detail_value(raw_value)
    if not category:
        fallback = _normalize_detail_value(f'{raw_category}: {raw_value}')
        if not fallback:
            return '', '', f'Unsupported detail category "{raw_category}".'
        return 'notes', fallback, ''
    if not value:
        return '', '', f'Empty value for "{category}" is not allowed.'
    return category, value, ''

def _normalize_detail_category(raw_category: str) -> str:
    key = re.sub(r'[\s\-]+', '_', str(raw_category or '').strip().lower())
    if not key:
        return ''
    return DETAIL_CATEGORY_ALIASES.get(key, '')

def _normalize_detail_value(raw_value: str) -> str:
    value = str(raw_value or '').strip()
    if not value:
        return ''
    if len(value) > MAX_DETAIL_VALUE_LENGTH:
        return value[:MAX_DETAIL_VALUE_LENGTH]
    return value

def ask_chart_inputs() -> tuple[str, dict[str, float]]:
    while True:
        print('What chart type would you like? (line, bar, pie, scatter)')
        chart_type = prompt_user('You: ').strip().lower()
        if chart_type in ('line', 'bar', 'pie', 'scatter'):
            break
        print('Invalid chart type. Please choose line, bar, pie, or scatter.')

    while True:
        print('Provide chart data as label:value pairs separated by commas.')
        print('Example: 2019:25, 2020:21, 2021:19')
        raw_data = prompt_user('You: ').strip()
        parsed, error = _parse_chart_data(raw_data)
        if parsed:
            return chart_type, parsed
        print(error)

def _parse_chart_data(raw_data: str) -> tuple[dict[str, float], str]:
    if not raw_data.strip():
        return {}, 'Chart data cannot be empty. Please use label:value pairs.'

    entries = [entry.strip() for entry in re.split(r'[,;\n]+', raw_data) if entry.strip()]
    parsed: dict[str, float] = {}
    seen_labels: set[str] = set()
    for entry in entries:
        if ':' not in entry:
            return {}, f'Invalid entry "{entry}". Please use label:value format.'
        label, value = entry.split(':', 1)
        label = label.strip()
        value = value.strip().replace('%', '')
        if not label:
            return {}, 'Label cannot be empty. Example: Jan:10, Feb:12'
        normalized_label = label.casefold()
        if normalized_label in seen_labels:
            return {}, f'Duplicate label "{label}" is not allowed. Use each label once.'
        try:
            parsed[label] = float(value)
        except ValueError:
            return {}, f'Value for "{label}" must be numeric.'
        seen_labels.add(normalized_label)
    return parsed, ''

def _chart_read_fallback(file_name: str) -> str:
    path = Path(file_name)
    if not path.exists():
        return f'Chart file "{file_name}" was not found.'
    size_bytes = path.stat().st_size
    return (
        f'Chart fallback summary:\n'
        f'- File: {path.resolve()}\n'
        f'- Format: {path.suffix.lower() or "unknown"}\n'
        f'- Size: {size_bytes} bytes\n'
        f'- Status: File exists, but automatic chart content extraction is not available in this environment.'
    )

def ask_formatting_options(ext: str) -> tuple[str, dict[str, object]]:
    if not ask_yes_no('Would you like to set formatting options (font, color, size, styles, alignment)? (yes/no)'):
        return '', {}

    options: dict[str, object] = {}
    selected_font = font_name(ext)
    if selected_font:
        options['font'] = selected_font

    selected_color = _ask_color()
    if selected_color:
        options['color'] = selected_color

    selected_size = _ask_font_size()
    if selected_size is not None:
        options['size'] = selected_size

    selected_styles = _ask_text_styles()
    if selected_styles:
        options['styles'] = selected_styles

    selected_alignment = _ask_alignment()
    if selected_alignment:
        options['alignment'] = selected_alignment

    style = selected_font if ext == 'txt' else ''
    return style, options

def _ask_color() -> str:
    print('Choose text color (name like red/blue or hex like #1E90FF). Press Enter to skip.')
    while True:
        value = prompt_user('You: ').strip()
        if not value:
            return ''
        lowered = value.lower()
        if re.fullmatch(r'#[0-9a-fA-F]{6}', value) or re.fullmatch(r'#[0-9a-fA-F]{3}', value):
            return lowered
        if lowered in ('black', 'white', 'red', 'green', 'blue', 'yellow', 'orange', 'purple', 'gray', 'grey', 'brown'):
            return lowered
        print('Invalid color. Use a simple color name or a hex color (example: #3366CC).')

def _ask_font_size() -> float | None:
    print('Choose font size in points (example: 12). Press Enter to skip.')
    while True:
        raw = prompt_user('You: ').strip()
        if not raw:
            return None
        try:
            size = float(raw)
        except ValueError:
            print('Invalid size. Please enter a number between 6 and 96.')
            continue
        if 6 <= size <= 96:
            return size
        print('Font size out of range. Please choose between 6 and 96.')

def _ask_text_styles() -> list[str]:
    print('Choose styles (comma separated): bold, italic, underline, uppercase, lowercase, title. Press Enter to skip.')
    supported = {'bold', 'italic', 'underline', 'uppercase', 'lowercase', 'title'}
    while True:
        raw = prompt_user('You: ').strip()
        if not raw:
            return []
        items = [item.strip().lower() for item in raw.split(',') if item.strip()]
        if not items:
            return []
        invalid = [item for item in items if item not in supported]
        if invalid:
            print(f'Unsupported style values: {", ".join(invalid)}. Please try again.')
            continue
        # Preserve order and remove duplicates
        unique_items = list(dict.fromkeys(items))
        case_styles = [item for item in unique_items if item in ('uppercase', 'lowercase', 'title')]
        if len(case_styles) > 1:
            print('Please choose only one of uppercase, lowercase, or title.')
            continue
        return unique_items

def _ask_alignment() -> str:
    print('Choose alignment: left, center, right, justify. Press Enter to skip.')
    supported = {'left', 'center', 'right', 'justify'}
    while True:
        raw = prompt_user('You: ').strip().lower()
        if not raw:
            return ''
        if raw in supported:
            return raw
        print('Invalid alignment. Please choose left, center, right, or justify.')

def _merge_format_options_from_details(
    format_options: dict[str, object],
    detail_items: list[tuple[str, str]],
) -> dict[str, object]:
    merged = dict(format_options or {})
    for category, value in detail_items:
        if category == 'fonts' and 'font' not in merged:
            candidate = value.strip().lower()
            if candidate in SUPPORTED_FONTS:
                merged['font'] = candidate
            continue

        if category == 'colors' and 'color' not in merged:
            candidate = value.strip().lower()
            if _is_valid_color_value(candidate):
                merged['color'] = candidate
            continue

        if category == 'sizes' and 'size' not in merged:
            parsed_size = _extract_size_value(value)
            if parsed_size is not None:
                merged['size'] = parsed_size
            continue

        if category == 'styles':
            parsed_styles = _extract_style_values(value)
            if parsed_styles:
                existing = merged.get('styles')
                if isinstance(existing, list):
                    combined = list(dict.fromkeys([*existing, *parsed_styles]))
                else:
                    combined = parsed_styles
                case_styles = [item for item in combined if item in ('uppercase', 'lowercase', 'title')]
                if len(case_styles) > 1:
                    first_case = case_styles[0]
                    combined = [item for item in combined if item not in ('uppercase', 'lowercase', 'title')]
                    combined.append(first_case)
                merged['styles'] = combined
            continue

        if category == 'alignments' and 'alignment' not in merged:
            candidate = value.strip().lower()
            if candidate in {'left', 'center', 'right', 'justify'}:
                merged['alignment'] = candidate
            continue
    return merged

def _is_valid_color_value(value: str) -> bool:
    candidate = str(value or '').strip().lower()
    if re.fullmatch(r'#[0-9a-f]{6}', candidate) or re.fullmatch(r'#[0-9a-f]{3}', candidate):
        return True
    return candidate in {'black', 'white', 'red', 'green', 'blue', 'yellow', 'orange', 'purple', 'gray', 'grey', 'brown'}

def _extract_size_value(value: str) -> float | None:
    match = re.search(r'(\d+(?:\.\d+)?)', str(value))
    if not match:
        return None
    try:
        size = float(match.group(1))
    except ValueError:
        return None
    if 6 <= size <= 96:
        return size
    return None

def _extract_style_values(value: str) -> list[str]:
    supported = {'bold', 'italic', 'underline', 'uppercase', 'lowercase', 'title'}
    pieces = [part.strip().lower() for part in re.split(r'[,/| ]+', str(value)) if part.strip()]
    return [part for part in list(dict.fromkeys(pieces)) if part in supported]

def _validate_file_name(name: str) -> tuple[bool, str]:
    if not name:
        return False, 'File name cannot be empty. Please include a name and extension (e.g., report.txt).'
    if ('/' in name or '\\' in name) and not _is_cloud_prefixed_name(name):
        return False, 'File name should not include path separators (/ or \\).'
    if name.endswith('.'):
        return False, 'File extension cannot end with a dot. Please provide a valid extension (e.g., .txt).'
    if '.' not in name:
        return False, 'Please include a file extension such as .txt or .docx.'

    base, ext = name.rsplit('.', 1)
    if not base.strip():
        return False, 'Please include characters before the extension.'
    if not ext.strip():
        return False, 'File extension cannot be empty. Please try again.'
    if not ext.isalnum():
        return False, 'File extension should only contain letters or numbers.'
    if len(ext) > 8:
        return False, 'File extension seems too long. Please double-check and try again.'

    return True, ''

def _is_cloud_prefixed_name(name: str) -> bool:
    if ':' not in name:
        return False
    prefix = name.split(':', 1)[0].strip().lower()
    return prefix in CLOUD_STORAGE_PREFIXES

def _default_ext_for_type(atype: str) -> str:
    if atype == 'txt':
        return 'txt'
    if atype == 'docx':
        return 'docx'
    if atype == 'xlsx':
        return 'xlsx'
    if atype == 'csv':
        return 'csv'
    if atype == 'pdf':
        return 'pdf'
    if atype == 'pptx':
        return 'pptx'
    if atype == 'markdown':
        return 'md'
    if atype == 'html':
        return 'html'
    if atype == 'code':
        return 'py'
    if atype == 'audio':
        return 'mp3'
    if atype == 'video':
        return 'mp4'
    if atype in ('image', 'chart'):
        return 'png'
    return 'txt'

def font_name(atype: str) -> str:
    while True:
        target = 'text output'
        if atype == 'txt':
            target = 'this txt file'
        elif atype == 'docx':
            target = 'this docx file'
        print(f'What font would you like to use for {target}? (press Enter to skip)')
        print()
        print('Supported fonts: Arial, Times New Roman, Calibri, Verdana, Georgia, Garamond, Comic Sans MS, Courier New, Tahoma, Helvetica')
        font = prompt_user('You: ').strip().lower()
        if not font:
            return ''
        if font in SUPPORTED_FONTS:
            return font
        print('Font name cannot be empty or misspelled (sorry!). Please try again.')
        continue

def file_name(action: str, atype: str) -> str:
    while True:
        if action == 'R':
            print('Please provide the full file name (including extension) you would like to read.')
        elif action == 'D':
            print('Please provide the full file name (including extension) you would like to delete.')
        elif action == 'A':
            print('What is the name of the file you want to append to? Include the extension (e.g., notes.txt).')
        else:
            print('What would you like to name the file? Please include an extension (e.g., draft.docx or todo.txt).')

        name_raw = prompt_user('You: ').strip()
        if not name_raw:
            print('File name cannot be empty. Please try again.')
            continue

        name = name_raw if '.' in name_raw else f'{name_raw}.{_default_ext_for_type(atype)}'

        valid, message = _validate_file_name(name)
        if valid:
            return name

        print(message)

def action_type() -> str:
    while True:
        print('What type of action would you like to perform? (read, write, append, delete)')
        action = prompt_user('You: ').strip().lower()

        if action in ('read', 'r'):
            return 'R'
        if action in ('write', 'w'):
            return 'W'
        if action in ('append', 'a'):
            return 'A'
        if action in ('delete', 'd'):
            return 'D'

        print('Invalid action. Please choose read, write, append, or delete.')

def ask_yes_no(message: str) -> bool:
    while True:
        print(message)
        answer = prompt_user('You: ').strip().lower()
        if answer in ('yes', 'y'):
            return True
        if answer in ('no', 'n', ''):
            return False
        print('Please answer with yes or no.')

def show_help():
    print('Available commands:')
    print('- Type your request to continue the current step.')
    print('- Type "help" at any prompt to see this message again.')
    print('- Type "quit" at any prompt to exit the application.')
    print('Supported types: txt, docx, xlsx, csv, pdf, pptx, markdown, html, code, image, chart, audio, video.')
    print('Cloud-style targets are supported with prefixes like "dropbox:report.docx" or "s3:folder/data.csv".')

def prompt_user(prompt: str) -> str:
    while True:
        try:
            response = input(prompt)
        except (EOFError, KeyboardInterrupt):
            print('\nThanks for using the File Generator Application! Goodbye!')
            sys.exit(0)

        lowered = response.strip().lower()
        if lowered in ('quit', 'exit', 'q'):
            print('Thanks for using the File Generator Application! Goodbye!')
            sys.exit(0)
        if lowered == 'help':
            show_help()
            continue
        return response

def file_type() -> str:
    while True:
        print('What type of file are you working with?')
        print('(txt, docx, xlsx, csv, pdf, pptx, markdown, html, code, image, chart, audio, video)')
        ftype = prompt_user('You: ').strip().lower()
        if ftype in ('txt', 'text', 'notepad'):
            return 'txt'
        if ftype in ('doc', 'docx', 'word'):
            return 'docx'
        if ftype in ('xls', 'xlsx', 'excel'):
            return 'xlsx'
        if ftype in ('csv',):
            return 'csv'
        if ftype in ('pdf',):
            return 'pdf'
        if ftype in ('ppt', 'pptx', 'powerpoint'):
            return 'pptx'
        if ftype in ('md', 'markdown'):
            return 'markdown'
        if ftype in ('html', 'htm'):
            return 'html'
        if ftype in (
            'code', 'programming', 'script',
            'py', 'js', 'java', 'c', 'cpp', 'rb', 'go', 'rs', 'swift', 'kt', 'ts', 'php',
            'json', 'xml', 'yml', 'yaml', 'toml', 'ini', 'sh', 'bat', 'ps1', 'css'
        ):
            return 'code'
        if ftype in ('image', 'img', 'picture', 'photo'):
            return 'image'
        if ftype in ('chart', 'graph', 'plot', 'pie', 'bar'):
            return 'chart'
        if ftype in ('audio', 'sound', 'music', 'mp3', 'wav', 'flac'):
            return 'audio'
        if ftype in ('video', 'movie', 'clip', 'mp4', 'avi', 'mkv', 'mov', 'wmv', 'webm'):
            return 'video'

        print('Invalid file type. Please choose one of the listed file types.')

def _supports_ai_processing(ext: str) -> bool:
    return ext in PROCESS_SUPPORTED_EXTENSIONS

def _is_failure(result: str) -> bool:
    if not isinstance(result, str):
        return False
    lower = result.lower()
    return (
        lower.startswith('error')
        or 'missing dependency' in lower
        or 'unsupported file type' in lower
        or 'unavailable' in lower
        or 'invalid action' in lower
    )



if __name__ == '__main__':
    main()
    print('Execution completed.')
    # if there are anythings I could do i would have obivously done it alresdy, but if you have any suggestions or feedback on how to improve the code, please l
# API KEY -- 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
# EOF end of file. sj this is the end of the file. do not add anything after this line.

#
