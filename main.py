import argparse

parser = argparse.ArgumentParser(description=r"""
Этот скрипт генерирует XML-файл с морфологической разметкой на основе excel-файла с выровненными предложениями.
Чтобы обработать файлы, помести их в папку texts\input. После завершения обработки эти файлы вместе с XML будут лежать в папке texts\output и будут называться так же, как и оригинальный файл, но с постфиксом _processed и другим расширением -- .xml.
Пожалуйста, укажи, каким языкам соответствуют какие названия колонок в excel-файлах, с помощью флагов --hye и --rus. 
Например, если в файле с выровненными предложениями предложения на русском лежат в столбце "ru", а предложения на (восточно-)армянском -- в столбце "hy", вызов скрипта будет выглядеть следующим образом:
>>> python main.py --hye hy --rus ru
"""
)

parser.add_argument("--hye", default='hy', help='Название столбца с предложениями на армянском')
parser.add_argument('--rus', default='ru', help='Название столбца с предложениями на русском')

args_ = parser.parse_args()

LANG_COLUMNS = {
    args_.hye: 'hye',
    args_.rus: 'rus'
}

if __name__ == '__main__':
    import os

    INPUT_DIR = os.path.join('texts', 'input')
    OUTPUT_DIR = os.path.join('texts', 'output')

    zips = list(filter(lambda x: x.endswith('.zip'), os.listdir(INPUT_DIR)))

    # если в папке input есть .zip архивы, то вытаскиваем из них все файлы в input, добавляя название архива в начало
    if zips:
        import zipfile
        for z in zips:
            with zipfile.ZipFile(os.path.join(INPUT_DIR, z), 'r') as zipdata:
                for zipinfo in zipdata.infolist():
                    # This will do the renaming
                    zipinfo.filename = z.replace('.zip', '') + '__' + zipinfo.filename
                    zipdata.extract(zipinfo, path=INPUT_DIR)

    files = list(filter(lambda x: x.endswith('.xlsx'), os.listdir(INPUT_DIR)))
    if files:

        from classes import XLSX2XML
        import datetime

        with open(os.path.join(OUTPUT_DIR, datetime.datetime.now().strftime('%Y%m%d%H%M')+'.txt'), 'w+', encoding='utf8') as stats_f:
            stats_f.write('\t'.join(['path', 'sent_num', 'hy_words', 'ru_words']) + '\n')
            for f in files:
                FILEPATH = os.path.join(INPUT_DIR, f)
                print(FILEPATH+'...')

                inst = XLSX2XML(filename=f, input_path=INPUT_DIR, output_path=OUTPUT_DIR, col_mapping=LANG_COLUMNS)
                inst.write_xml()

                stats_f.write('\t'.join(inst.stats) + '\n')

