rem requires three filenames 1 wav 2 txt 3 srt
python -m aeneas.tools.execute_task %1 %2 "task_language=fa|is_text_type=plain|os_task_file_format=srt" %3
pause