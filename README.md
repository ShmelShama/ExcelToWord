# ExcelToWord
Выполненное тестовое задание на позицию Программист-разработчик

Задание реализовано в виде приложения WPF
1.	Приложение загружает данные из файла Excel формата
2.	По загруженным данным создается отчет в формате .doc
3.	Приложение открывает созданный отчет


## Packages

| Package   | URL/CLI                                     |
| -------- | ---------------------------------------- | 
| **.NET Framework 4.8**    |https://dotnet.microsoft.com/ru-ru/download/dotnet-framework/net48               |
| **Microsoft.Office.Interop.Word**   |dotnet add package Microsoft.Office.Interop.Word --version 15.0.4797.1004                            | 
| **Microsoft.Office.Interop.Excel**    | dotnet add package Microsoft.Office.Interop.Excel --version 15.0.4795.1001   | 



## Интерфейс
![image](https://github.com/user-attachments/assets/79cac04d-f7e6-46c5-a65b-91550989722b)


## Работа с программой
1.	Выберите Excel файл из которого будут загружены данные.

**Названия листов и наименования столбцов должны быть такими же как в указанном ниже примере! Иначе программа посчитает данные недостоверными или неполными.**

2.	Укажите папку, где будет создан отчет в формате .doc
3.	Нажмите Приступить.

### Пример входных данных в Excel

Файл должен содержать три листа: Сотрудники, Задачи, Отделы (порядок не важен)

Пример данных по сотрудникам:

![image](https://github.com/user-attachments/assets/00c74df5-487e-440c-82ca-49af95fad970)

Пример данных по отделам:

![image](https://github.com/user-attachments/assets/6d0359da-bbef-4356-99fa-09957b10b109)


Пример данных по задачам:

![image](https://github.com/user-attachments/assets/3a73e49a-34ef-4780-b43e-3a4bc30720b7)

