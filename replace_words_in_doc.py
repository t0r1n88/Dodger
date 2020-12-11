import docx
import os

# Список для сохранения имен документов

# Получаем текущую папку и добавляем еще одну подпапку
folder = os.getcwd() + '\data'


def get_list_files(path, extension='docx'):
    """
    Функция для получения списка всех  файлов внутри определенной папки с определенным разрешением
    :param path:Путь к папке
    :return:список файлов
    """
    paths = []
    # Рекурсивно перебираем папки и файлы находящиеся в папке path
    for root, dirs, files in os.walk(path):
        # Перебираем найденные файлы чтобы отобрать файлы с нужным расширением
        for file in files:
            if file.endswith(extension) and not file.startswith('~'):
                paths.append(os.path.join(root,file))
    return paths
# Получаем список файлов в которых нужно произвести замену
files = get_list_files(folder)

for file in files:
    # Получаем  объект файла
    doc = docx.Document(file)
    for paragraph in doc.paragraphs:
        if 'Будаев О.Т' in paragraph.text:
            print('Найдены отклонения')
            with open('mistakes.txt','a',encoding='utf8') as f:
                f.write(f'{file}\n')
