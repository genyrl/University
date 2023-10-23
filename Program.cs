using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Data
{

    internal class Program
    {
        static void Main(string[] args)
        {
            Help();
            List<Student> students = new List<Student>()
            {
             new Student("Висковатов",1,2050),
             new Student("Стасян",3,1000),
             new Student("Ленин",2, 5000)
            };
            bool exit = true;
            do
            {
                string[] com = Console.ReadLine().Trim().Split();
                Console.WriteLine();
                switch (com[0])
                {
                    case "прочитать":
                        if (com.Length == 2 && com[1].EndsWith(".xlsx"))
                            ReadExsel(students, com[1]);
                        else
                            Console.WriteLine("Неверно указан путь к файлу");
                        break;
        
                    case "очистить":
                        Console.Clear();
                        break;
                    case "изменить":
                        int id, curse, fee;

                        if (com.Length == 5 && int.TryParse(com[1], out id) && id - 1 >= 0 && id - 1 < students.Count && int.TryParse(com[3], out curse) && int.TryParse(com[4], out fee))
                        {
                            students[id - 1].Edit(com[2], curse, fee);
                        }

                        else
                            Console.WriteLine("Неверно введена команда!");
                        break;
                    case "удалить":

                        if (int.TryParse(com[1], out id) != false && students.Count > id - 1 && (id - 1) >= 0)
                            Delete(students, id - 1);
                        else if (students.Find(x => x.LastName == com[1]) != null)
                        {
                            Delete(students, com[1]);
                        }
                        break;
                    case "добавить":

                        if (com.Length == 5 && int.TryParse(com[1], out id) && id - 1 >= 0 && id -1 <= students.Count && int.TryParse(com[3], out curse) && int.TryParse(com[4], out fee))
                            if (students.Count < 300)
                                Add(students, id - 1, new Student(com[2], curse, fee));
                            else
                                Console.WriteLine("Максимальное кол-во данных");
                        else
                            Console.WriteLine("Неверно введена команда!");
                        break;
                    case "таблица":

                        CreateExsel(students);
                        break;

                    case "вывод":
                        if (com.Length == 2 && (int.TryParse(com[1], out id) != false && students.Count > id - 1 && (id - 1) >= 0))
                            PrintCon(students, id - 1);
                        else if (com.Length == 1)
                            PrintCon(students);
                        else Console.WriteLine("Неверно указан индекс или команда!");
                        break;

                    case "сорт":

                        if (com.Length == 3 && (com[1] =="имя" || com[1] == "курс" || com[1] == "стипендия") && (com[2] == "убывание" || com[2] == "возрастание"))
                          students =  Sortby(students, com[1], com[2]);
                        else
                            Console.WriteLine("Неверно написана команда Сорт!");
                        break;
                    case "помощь":

                        Help();
                        break;

                    case "":

                        exit = false;
                        break;

                    default:

                        Console.WriteLine("Неверно указана команда");
                        break;


                }
            } while (exit);
        }

        static void CreateExsel(IEnumerable<Student> students)
        {
            var exselApp = new Application();

            exselApp.Visible = true;

            exselApp.Workbooks.Add();
            _Worksheet worksheet = (Worksheet)exselApp.ActiveSheet;

            worksheet.Cells[1, "A"] = "Фамилия";
            worksheet.Cells[1, "B"] = "Номер Курса";
            worksheet.Cells[1, "C"] = "Стипендия";

            var row = 1;
            foreach (Student student in students)
            {
                row++;

                worksheet.Cells[row, "A"] = student.LastName;
                worksheet.Cells[row, "B"] = student.Curse;
                worksheet.Cells[row, "C"] = student.Fee;

            }
            for (int i = 1; i < 4; i++)
            {
                worksheet.Columns[i].AutoFit();
            }

        }
        static void ReadExsel(List<Student> students, string dir)
        {

            Application excelApp = new Application();

            Workbook excelBook = excelApp.Workbooks.Open(dir);
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            students.Clear();
            for (int i = 2; i <= rows; i++)

            {
                if (excelRange.Cells[i, "A"] != null && excelRange.Cells[i, "A"].Value2 != null)
                    if (students.Count < 300)
                        students.Add(new Student(excelRange.Cells[i, "A"].Value2.ToString(), (int)excelRange.Cells[i, "B"].Value2, (int)excelRange.Cells[i, "C"].Value2));
                    else
                    {
                        Console.WriteLine("Прочитаны первые 300 столбцов");
                        break;
                    }
            }
          
            excelApp.Quit();
        }

        static void PrintCon(List<Student> students, int id = -1)
        {
            if (id == -1)
            {
                int maxSim = 0;
                string st = "";
                foreach (Student student in students)
                {
                    maxSim = Math.Max(maxSim, student.LastName.Length);
                }
                for (int i = 0; i < maxSim - 7; i++)
                {
                    st += " ";
                }
                Console.WriteLine($"Фамилия {st} Курс  Стипендия");
                int a = 8 + st.Length;
                foreach (Student student in students)
                {
                    st = "";
                    for (int i = 0; i < a - student.LastName.Length; i++)
                    {
                        st += " ";
                    }
                    Console.WriteLine($"{student.LastName} {st} {student.Curse}      {student.Fee}");
                }
            }
            else
            {
                Console.WriteLine(students[id].ToString());
            }
        }
        static List<Student>  Sortby(List<Student> students, string sort, string forward)
        {
           
            switch (sort)
            {
                case "имя":
                    students.Sort(new SortByName());
                    break;
                case "курс":
                    students.Sort(new SortByCurse());
                    break;
                case "стипендия":
                    students.Sort(new SortByFee());
                    break;
                
            }
            if (forward == "убывание")
                students.Reverse();
            PrintCon(students);
            return students;


        }

        static void Add(List<Student> students, int id, Student nw)
        {

            students.Insert(id, nw);


        }
        static void Delete(List<Student> students, int id)
        {
            students.RemoveAt(id);
        }
        static void Delete(List<Student> students, string name)
        {
            students.Remove(students.Find(x => x.LastName == name));
        }

        static void Help()
        {
            Console.WriteLine("Команды: " +
                "\nтаблица - Создать таблицу" +
                "\nвывод {номер| } - Вывести на консоль номер студента или всю таблицу" +
                "\nсорт (имя|курс|стипендия) (убывание|возрастание) - Сортировать по (Имя|Курс|Стипендия) на (Убывание|Возрастание) и Вывести на консоль" +
                "\nпомощь - список команд" +
                "\nдобавить {строка(число)} {фамилия} {курс(число)} {стипендия(число)}- Добавить студента в таблицу" +
                "\nудалить {строка(число)|фамилия(при одинаковых удалится верхняя)}" +
                "\nизменить {столбец(число)} {фамилия} {курс(число)} {стипендия(число)}" +
                "\nочистить - очистить консоль" +
                "\nпрочитать {путь к файлу} - чтение таблицы эксель(первый три столбца = фамилия,курс,стипендия) ");
        }
    }

    public class Student
    {

        public string LastName { get; set; }
        public int Curse { get; set; }
        public double Fee { get; set; }

        public Student(string lastName, int curse, int fee)
        {
            LastName = lastName;
            Curse = curse;
            Fee = fee;

        }

        public override string ToString() => $"{LastName} {Curse} {Fee}";
        public void Edit(string name, int course, double fee)
        {
            LastName = name;
            Curse = course;
            Fee = fee;
        }
        public void Edit(string name, int course)
        {
            Edit(name, course, Fee);
        }
    }

    public class SortByName: IComparer<Student>
    {
    
        int IComparer<Student>.Compare(Student x, Student y)
        {
            return String.Compare(x.LastName, y.LastName);
        }

    }
    public class SortByCurse: IComparer<Student>
    {
        int IComparer<Student>.Compare(Student x, Student y)
        {
            if (x.Curse > y.Curse) return 1;
            else if (x.Curse<y.Curse) return -1;
            else return 0;
        }
    }
    public class SortByFee : IComparer<Student>
    {
        int IComparer<Student>.Compare(Student x, Student y)
        {
            if(x.Fee > y.Fee) return 1;
            else if(x.Fee < y.Fee) return -1; 
            else return 0;
        }
    }


}
