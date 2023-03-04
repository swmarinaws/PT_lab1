import random
from russian_names import RussianNames
import pandas as pd
import time
import copy

list_faculties = ['Биологический', 'Богословский','Географический', 'Геологический', 'Журналистики', 'Информационный', 'Исторический', 'Кибернетики', 'Математический', 'Механический', 'Политологический', 'Психологический', 'Радиотехнический', 'Социологический', 'Управления', 'Физический', 'Филологический', 'Философский', 'Химический', 'Художественно-графический', 'Экономический', 'Юридический']
list_ranks = ['Доцент', 'Профессор']
list_degrees = ['Кандидат наук', 'Доктор наук']
sizes = [100, 250, 500, 1000, 5000, 10000, 100000]

"""Время сортировки для каждого алгоритма"""
time_bubble = []
time_quick = []
time_merge = []

def generate(n):
    """Генерирование n данных"""
    final_dict = {}
    names, surnames, patronymics, faculties, ranks, degrees = [], [], [], [], [], []
    for i in range(n):
        full_name = RussianNames().get_person().split()
        names.append(full_name[0])
        patronymics.append(full_name[1])
        surnames.append(full_name[2])
        faculties.append(random.choice(list_faculties))
        ranks.append(random.choice(list_ranks))
        degrees.append(random.choice(list_degrees))

    final_dict['Фамилия'] = surnames
    final_dict['Имя'] = names
    final_dict['Отчество'] = patronymics
    final_dict['Факультет'] = faculties
    final_dict['Учёная степень'] = ranks
    final_dict['Учёное звание'] = degrees
    return final_dict

def bubble_sort(l):
    """Сортировка пузрьком"""
    for i in range(len(l)):
        for j in range(len(l) - 1, i, -1):
            if l[j - 1] > l[j]:
                l[j - 1], l[j] = l[j], l[j - 1]

def quick_sort(l, fst, lst):
    """Быстрая сортировка"""
    if fst >= lst: return

    i, j = fst, lst
    pivot = l[fst + (lst - fst) // 2]

    while i <= j:
        while l[i] < pivot: i += 1
        while l[j] > pivot: j -= 1
        if i <= j:
            l[i], l[j] = l[j], l[i]
            i += 1
            j -= 1

    quick_sort(l, fst, j)
    quick_sort(l, i, lst)

def merge(l, low, mid, high):
    """Функция, которая записывает отсортированный подмассив в оригинальный массив"""
    b = [None] * (high + 1 - low)

    h = 0
    i = low
    j = mid + 1

    while i <= mid and j <= high:
        if l[i] <= l[j]:
            b[h] = l[i]
            i += 1
        else:
            b[h] = l[j]
            j += 1
        h += 1

    if i > mid:
        for k in range(j, high + 1):
            b[h] = l[k]
            h += 1
    else:
        for k in range(i, mid + 1):
            b[h] = l[k]
            h += 1

    for k in range(0, high - low + 1):
        l[low + k] = b[k]

def merge_sort(l, low, high):
    """Сортировка слиянием"""
    if low < high:
        mid = (low + high) // 2
        merge_sort(l, low, mid)
        merge_sort(l, mid + 1, high)
        merge(l, low, mid, high)

class Teacher:
    """Класс для описания объекта преподавателя"""
    """Объект включает в себя: фамилию, имя, отчество, факультет, учёное звание, учёную степень"""
    def __init__(self, surname, name, patronymic, faculty, rank, degree):
        self.surname = surname
        self.name = name
        self.patronymic = patronymic
        self.faculty = faculty
        self.rank = rank
        self.degree = degree

    def __gt__(self, other):
        """Перегрузка оператора >"""
        if self.faculty != other.faculty:
            return self.faculty > other.faculty
        if self.surname != other.surname:
            return self.surname > other.surname
        if self.name != other.name:
            return self.name > other.name
        if self.patronymic != other.patronymic:
            return self.patronymic > other.patronymic
        if self.degree != other.degree:
            return self.degree > other.degree
        return self.rank > other.rank

    def __lt__(self, other):
        """Перегрузка оператора <"""
        if self.faculty != other.faculty:
            return self.faculty < other.faculty
        if self.surname != other.surname:
            return self.surname < other.surname
        if self.name != other.name:
            return self.name < other.name
        if self.patronymic != other.patronymic:
            return self.patronymic < other.patronymic
        if self.degree != other.degree:
            return self.degree < other.degree
        return self.rank < other.rank

    def __ge__(self, other):
        """Перегрузка оператора >="""
        if self.faculty != other.faculty:
            return self.faculty >= other.faculty
        if self.surname != other.surname:
            return self.surname >= other.surname
        if self.name != other.name:
            return self.name >= other.name
        if self.patronymic != other.patronymic:
            return self.patronymic >= other.patronymic
        if self.degree != other.degree:
            return self.degree >= other.degree
        return self.rank >= other.rank

    def __le__(self, other):
        """Перегрузка оператора <="""
        if self.faculty != other.faculty:
            return self.faculty <= other.faculty
        if self.surname != other.surname:
            return self.surname <= other.surname
        if self.name != other.name:
            return self.name <= other.name
        if self.patronymic != other.patronymic:
            return self.patronymic <= other.patronymic
        if self.degree != other.degree:
            return self.degree <= other.degree
        return self.rank <= other.rank

"""Запись сгенерированных данных в файл MS Excel"""
with pd.ExcelWriter("./sets.xlsx") as writer:
    for i in sizes:
        pd.DataFrame(generate(i)).to_excel(writer, sheet_name=f"{i}", index=False)

"""Считывание входных данных из файла MS Excel и запись в словарь"""
teachers = {}
for i in sizes:
    temp = pd.read_excel('./sets.xlsx', sheet_name=f"{i}").to_dict('records')
    teachers[i] = [Teacher(t['Фамилия'], t['Имя'], t['Отчество'], t['Факультет'], t['Учёная степень'], t['Учёное звание']) for t in temp]

"""Реализация сортировок входных данных"""
for i in sizes:
    sorted_arrays = []

    sorted_arr_bubble = copy.deepcopy(teachers[i])
    start = time.time()
    bubble_sort(sorted_arr_bubble)
    end = time.time() - start
    time_bubble.append(end)
    sorted_arrays.append(sorted_arr_bubble)

    sorted_arr_quick = copy.deepcopy(teachers[i])
    start = time.time()
    quick_sort(sorted_arr_quick, 0, len(sorted_arr_quick) - 1)
    end = time.time() - start
    time_quick.append(end)
    sorted_arrays.append(sorted_arr_quick)

    sorted_arr_merge = copy.deepcopy(teachers[i])
    start = time.time()
    merge_sort(sorted_arr_merge, 0, len(sorted_arr_merge) - 1)
    end = time.time() - start
    time_merge.append(end)
    sorted_arrays.append(sorted_arr_merge)

    for j in range(len(sorted_arrays)):
        final_dict = {}
        names, surnames, patronymics, faculties, ranks, degrees = [], [], [], [], [], []

        for k in sorted_arrays[j]:
            names.append(k.name)
            surnames.append(k.surname)
            patronymics.append(k.patronymic)
            faculties.append(k.faculty)
            ranks.append(k.rank)
            degrees.append(k.degree)

        final_dict['Фамилия'] = surnames
        final_dict['Имя'] = names
        final_dict['Отчество'] = patronymics
        final_dict['Факультет'] = faculties
        final_dict['Учёная степень'] = ranks
        final_dict['Учёное звание'] = degrees

        """Запись в файл с соответствующим методом сортировки"""
        if j == 0:
            file_name = "./sets_bubble.xlsx"
        elif j == 1:
            file_name = "./sets_quick.xlsx"
        else:
            file_name = "./sets_merge.xlsx"

        """Проверка: если первый набор данных, то создание файла, иначе запись в существующий"""
        if i == 100:
            with pd.ExcelWriter(file_name, engine="openpyxl", mode='w') as writer:
                pd.DataFrame(final_dict).to_excel(writer, sheet_name=f"{i}", index=False)
        else:
            with pd.ExcelWriter(file_name, engine="openpyxl", mode='a') as writer:
                pd.DataFrame(final_dict).to_excel(writer, sheet_name=f"{i}", index=False)

"""Вывод времени, потраченного на каждую сортировку"""
print(f'Bubble_time = {time_bubble}')
print(f'Quick_time = {time_quick}')
print(f'Merge_time = {time_merge}')