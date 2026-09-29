import json
from ollama import chat

from pathlib import Path

from analyzer_file import MetaAnalyzerFile

path_file = "C:\\Users\\ilyam\\Desktop\\Пр-во+экспрессы ноябр-дек21г.xlsx"

meta_analyzer = MetaAnalyzerFile.MetaAnalyzerFile(path_file)


def send_request_to_ollama(request: str):
    system_prompt = '''
        Ты — анализатор Excel-файлов.
        Твоя задача — анализировать структуру Excel-файла на основе данных профайлера.
        
        СТРОГОЕ ПРАВИЛО:
        - Весь ответ должен быть написан ТОЛЬКО на русском языке.
        - Не используй английский язык в ответе.
        - Названия ключей JSON, имена листов, адреса ячеек и названия типов данных можешь оставлять как есть.
        - Все объяснения, выводы и рекомендации пиши на русском языке.
    '''

    user_prompt = f'''
        Проанализируй структуру Excel-файла и найди потенциальные проблемы.
        Данные профайлера:
        {request}
    '''
    
    response = chat(
        model='qwen3:8b',
        messages=[
            {
                'role': 'system',
                'content': system_prompt
            },
            {
                'role': 'user',
                'content': user_prompt
            }
        ]
    )
    return response.message.content


def save_response(response: str):
    base_dir = Path(__file__).resolve().parent
    save_response_dir = base_dir / 'save_response'

    save_response_dir.mkdir(exist_ok=True)

    file_save = save_response_dir / 'save_response.txt'
    path_file = Path(file_save)

    with open(path_file, 'w', encoding='utf-8') as file:
        file.write(response)


def main():
    request_data = meta_analyzer.create_profiler_file()
    request = json.dumps(request_data, ensure_ascii=False, indent=4)  # -> готовый запрос к Ollama

    print(f"Размер JSON: {len(request.encode('utf-8')) / 1024:.2f} KB")

    response = send_request_to_ollama(request)
    print(response)

    save_response(response)


if __name__ == '__main__':
    main()
