# Подключаем необходимые модули
from tkinter import *
from tkinter.messagebox import *
import openpyxl as xl
from random import randint


# Определение позиции последнего символа слова
def last_symb(s):
    # Строка из "плохих" символов
    bad = "йъьы"

    # Проверка слова на наличие "плохого" символа в слове
    for b in bad:
        if s[-1] == b:
            return -2
    return -1


# Окно правил игры
def game_rules():
    # Создание окна
    win = Toplevel()

    # Название окна
    win.title('Правила')

    # Запрет на изменение размера окна
    win.resizable(0, 0)

    # Устанавливаем размер окна и позиционируем по центру
    center_position(win, 600, 340)

    # Иконка окна
    win.iconbitmap('gamepad_icon.ico')

    # Нельзя закрыть родительское окно, пока открыто дочернее
    win.grab_set()

    # Фокусируем на этом окне
    win.focus_force()

    # Создаем метки, оформляем и упаковываем их
    r1 = Label(win, text="Правила игры:", justify='left', font="Arial 14", fg="red4")
    r2 = Label(win, text=""" 
    1. Назвать слово, начинающееся на последнюю букву другого слова.\n
    2. Слово должно быть не меньше двух букв.\n
    3. Если слово заканчивается на "й", "ъ", "ь" и "ы", то назвать слово на предпоследнюю букву.\n
    4. Слова не должны повторяться.\n
    5. Если у компьютера закончится словарный запас, то Вы выиграете.\n
    6. В любой момент можно сдаться.\n
    7. Наслаждайтесь процессом игры.\n
                            """, justify='left', font="Arial 10")
    r1.pack()
    r2.pack()

    # Создаем кнопку, оформляем и упаковываем
    b = Button(win, text="Понятно", command=win.destroy, font="Arial 13", bg="red4", fg="white smoke")
    b.pack()


# Окно игрового процесса
def game_process(event):

    # Деактивация кнопки при несоответствии правилам
    def active_button(*arg):

        # Берем значение text из Entry
        a = for_check.get()

        # Проверка на соответствие введенного текста правилам
        if a != "" and a.isalpha() and len(a) >= 2:

            # Кнопка активирована
            btn.config(state='normal')
        else:

            # Кнопка деактивирована
            btn.config(state='disabled')

    # Критический анализ слова игрока и получение слова компьютера
    def get_pc_word(user_word, pc_last_symbol):

        # Подключаемся к таблице слов
        wb = xl.load_workbook(filename="vocabulary.xlsx")

        # Обратимся к первому листу
        wb.active = 0
        sheet = wb.active

        # Берем столбец (тему игры)
        th = theme_var.get()

        # Увеличиваем регистр первой буквы слова игрока
        user_word = user_word.capitalize()

        # Увеличиваем регист последней буквы слова компьютера
        pc_last_symbol = pc_last_symbol.upper()

        # Если введенное слово не пустое
        if len(user_word) != 0:

            # Если буквы совпадают
            if user_word[0] == pc_last_symbol:

                # Если слова игрока нет в списке истории слов
                if user_word not in history_of_words:

                    # Добавляем слово в список история слов
                    history_of_words.append(user_word)

                    # Определяем последняя или предпоследняя буква
                    l_user = last_symb(user_word)

                    # Находим слово компьютера в таблице
                    for w in range(1, 58):
                        pc_word = sheet[th + str(w)].value
                        print("PC ", pc_word)

                        # Если первая буква слова компьютера равна последней букве слова пользователя
                        if pc_word[0] == user_word[l_user].upper():

                            # Проверка этого слова пк в список истории слов
                            if pc_word not in history_of_words:

                                # Добавляем слово компьютера в историю слов
                                history_of_words.append(pc_word)

                                # Присваиваем значение слова компьютера метапеременной
                                t.set(pc_word)
                                return pc_word

                    # Если компьютер слова не нашел, то игрок выиграл
                    btn2.config(state='disabled')
                    btn.config(state='disabled')
                    return win_or_lose(1, win_game)
                else:
                    # Игрок проиграл, кнопки деактивировались
                    btn2.config(state='disabled')
                    btn.config(state='disabled')
                    return win_or_lose(0, win_game)
            else:
                # Игрок проиграл, кнопки деактивировались
                btn2.config(state='disabled')
                btn.config(state='disabled')
                return win_or_lose(0, win_game)
        else:
            # Взять любое слово из темы
            pc_word = sheet[th + str(randint(1, 58))].value

            # Добавить это слово в список история слов
            history_of_words.append(pc_word)

            # Присвоить компьютерное слово метапеременной
            t.set(pc_word)
            return pc_word

    # Скрываем корневое окно
    root.withdraw()

    # Создаем окно игры
    win_game = Toplevel()

    # Запрещаем изменять размер окна
    win_game.resizable(0, 0)

    # Если игрок выйдет из игры, то показать диалоговое окно
    win_game.protocol("WM_DELETE_WINDOW", ask_exit_from_game)

    # Устанавливаем размер окна и центрируем в центре
    center_position(win_game, 500, 300)

    # Фокусируем на этом окне
    win_game.focus_force()

    # Иконка окна
    win_game.iconbitmap('gamepad_icon.ico')

    # Меню окна
    menu(win_game)

    # Создание и оформление метки
    lab = Label(win_game, text="Я называю", font="Arial 14")
    lab.pack(pady=5)

    # Создадим метапеременную слова компьютера
    t = StringVar()
    t.set(get_pc_word("", ""))

    # Вывод слова компьютера на экран
    lab2 = Label(win_game, textvariable=t, font="Arial 20", fg="red4")
    lab2.pack(pady=15)

    # Создание и оформление метки
    l1 = Label(win_game, text="Назовите слово:", font="Arial")
    l1.pack()

    # Создание метапеременной поля ввода слова
    for_check = StringVar()

    # Создание поля ввода слова и его оформление
    usr_input = Entry(win_game, textvariable=for_check, font="Arial", fg="red4")
    usr_input.pack()

    # Проверяем активированность кнопки по значению поля ввода
    for_check.trace("w", active_button)

    # Создаем и оформляем кнопки
    btn = Button(win_game, text="Назвать", state='disabled',
                 command=lambda: get_pc_word(usr_input.get(), t.get()[last_symb(t.get())]),
                 font="Arial 13", bg="red4", fg="white smoke")
    btn.pack(pady=9)

    btn2 = Button(win_game, text="Я сдаюсь", command=lambda: win_or_lose(0, win_game),
                  font="Arial", fg="red4")
    btn2.pack(pady=20)


# Формула позиционирования окна по центру экрана
def center_position(win, w, h):
    win.geometry(f"{w}x{h}+{(win.winfo_screenwidth() - w) // 2}+{(win.winfo_screenheight() - h) // 2}")


# Диалоговое окно "Спросить выйти ли из игры?"
def ask_exit_from_game():
    answer = askquestion("Выход из игры", "Вы действительно хотите выйти из игры?")

    # Если пользователь нажал "Да"
    if answer == "yes":

        # Выход из программы
        quit()


# Диалоговое окно "Начать игру заново?"
def ask_restart_game(wind):
    answer = askquestion("Начать заново", "Вы хотите начать игру заново?")

    # Если пользователь нажал "да"
    if answer == "yes":

        # Очистить историю игры
        history_of_words.clear()

        # Если это окно корневое
        if wind == root:

            # Ничего не делаем
            pass

        # Если окно дочернее
        else:

            # Уничтожаем дочернее окно
            wind.destroy()

            # Делаем корневое окно видимым
            root.deiconify()


# Диалоговые окна "Выиграл" и "Проиграл"
def win_or_lose(result, wind):

    # Очищаем историю игры
    history_of_words.clear()

    # Если игрок выиграл
    if result == 1:

        # Показать диалоговое окно
        msg = showinfo('Победа!', 'Вы выиграли!')

        # Если игрок нажал "ok"
        if msg == 'ok':

            # Уничтожаем дочернее окно
            wind.destroy()

            # Делаем корневое окно видимым
            root.deiconify()
    else:

        # Показать диалоговое окно
        msg = showinfo('Провал', 'Вы проиграли')

        # Если пользователь нажал "ok"
        if msg == 'ok':

            # Уничтожаем дочернее окно
            wind.destroy()

            # Делаем корневое окно видимым
            root.deiconify()


# Выпадающее меню
def menu(win):
    game_menu = Menu(win)
    win.configure(menu=game_menu)
    game_menu.add_cascade(label="Начать заново", command=lambda: ask_restart_game(win))
    game_menu.add_cascade(label="Правила игры", command=game_rules)
    game_menu.add_cascade(label="Выход из игры", command=ask_exit_from_game)


# История слов
history_of_words = []

# Названия тем
themes = {
    'A': 'Города',
    'B': 'Профессии',
    'C': 'Животные',
}

# Создание и настройка начального окна
root = Tk()

# Название окна
root.title('Игра в слова')

# Запрещаем изменение размера окна
root.resizable(0, 0)

# Определяем размеры экрана и позиционируем по центру
center_position(root, 500, 300)

# Иконка окна
root.iconbitmap('gamepad_icon.ico')

# Если игрок выйдет из игры, то показать диалоговое окно
root.protocol("WM_DELETE_WINDOW", ask_exit_from_game)

# Выпадающее меню
menu(root)

# Открыть окно правил игры через 0,1 секунду
root.after(100, game_rules)

# Создаем и оформляем метку
name_label = Label(root, text="Игра в слова", fg="red4", font="Arial 25")
name_label.pack(pady="45")

# Создаем фрейм "Выбор темы"
frame_themes = LabelFrame(root, text='Выберите тему для игры', font="Arial 10")
frame_themes.pack()

# Создаем метапеременную темы
theme_var = StringVar(value="A")

# Создаем радиокнопки
for theme in sorted(themes):
    Radiobutton(frame_themes, text=themes[theme],
                variable=theme_var,
                value=theme,
                font="Arial"
                ).pack(side='left')

# Создаем и оформляем кнопку
btn_start = Button(root, text="Начать игру", font="Arial 13", bg="red4", fg="white smoke")
btn_start.pack(pady="35")

# Нажатие кнопки приведет к процессу игры
btn_start.bind("<Button-1>", game_process)

# Зациклим окно
root.mainloop()