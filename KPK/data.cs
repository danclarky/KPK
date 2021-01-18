using Npgsql;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace KPK
{
    class data
    {
        public partial class options : data
        {
            public static string[] workmen = { "оповещение о платеже", "оповещение заёмщика", "оповещение поручителя(залогодателя)", "оповещение других лиц", "оповещение семьи", "оповещение места работы", "личная беседа", "комментарий" };
            public static string[] worksb = { "оповещение заёмщика", "оповещение поручителя(залогодателя)", "оповещение других лиц", "оповещение семьи", "оповещение места работы", "личная беседа", "выезд по месту жительства", "выезд по месту работы", "выезд к поручителю", "претензия", "передано юристу на подачу в суд", "комментарий", "заявление на судебный приказ", "исковое заявление", "назначено рассмотрение в суде/получена повестка", "посещение судебного заседания", "получено судебное решение", "получен исполнительный лист","получен судебный приказ", "отмена судебного решения", "смерть клиента", "заявление в ССП", "Ознакомление с материалами ИП", "заявление в ПФР, банк, работодателю", "жалоба в ССП", "иное", "ИП возбуждено", "ИП прекращено", "приостановление подачи в суд"};
            public static string[] urist = { "оповещение заёмщика", "оповещение поручителя(залогодателя)", "оповещение других лиц", "оповещение семьи", "оповещение места работы", "личная беседа", "выезд по месту жительства", "выезд по месту работы", "выезд к поручителю", "претензия", "комментарий", "заявление на судебный приказ", "исковое заявление", "назначено рассмотрение в суде/получена повестка", "посещение судебного заседания", "получено судебное решение", "получен исполнительный лист","получен судебный приказ", "отмена судебного решения", "смерть клиента", "заявление в ССП", "Ознакомление с материалами ИП", "заявление в ПФР, банк, работодателю", "жалоба в ССП", "иное", "ИП возбуждено", "ИП прекращено","Банкротство" };
            public static string[] zayavka = { "Звонок", "Личная беседа", "Оформление", "Отказ" };
            public static string[] ringing = { "Недоступен", "Занято", "Придет", "Заключил договор", "Отказ", "Есть активный заем", "Просит перезвонить" };
        }

        public static string[] users()
        {
            string[] usersar;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from Пользователи where Роль='Менеджер'";
                    object value = cmd.ExecuteScalar();
                    usersar = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select ФИО from Пользователи where Роль='Менеджер' order by ФИО";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            usersar[i] = reader["ФИО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return usersar;
        }

        public static string[] stringconnect()
        {
            string[] usersar;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    usersar = new string[6];
                    cmd.CommandText = "select Строка from ПодключениеБаз order by Код";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            usersar[i] = reader["Строка"].ToString();
                            i++;
                        }
                    }
                }
            }
            return usersar;
        }

        public static string[] DO()
        {
            string[] DOO;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from ДО where Актив='+'";
                    object value = cmd.ExecuteScalar();
                     DOO = new string[Convert.ToInt32(value)];
                     cmd.CommandText = "select ДО from ДО where Актив='+' order by ДО";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DOO[i] = reader["ДО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return DOO;
        }
        public static string[] template()
        {
            string[] template;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from ШаблоныЗаймов";
                    object value = cmd.ExecuteScalar();
                    template = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select Шаблон from ШаблоныЗаймов order by Шаблон";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            template[i] = reader["Шаблон"].ToString();
                            i++;
                        }
                    }
                }
            }
            return template;
        }
        public static string[] DOactual()
        {
            string[] DOO;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from ДО where Актив='+'";
                    object value = cmd.ExecuteScalar();
                    DOO = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select ДО from ДО where Актив='+' order by ДО ";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DOO[i] = reader["ДО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return DOO;
        }
        public static string[] DOanal()
        {
            string[] DOO;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from ДО where аналитика='+'";
                    object value = cmd.ExecuteScalar();
                    DOO = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select ДО from ДО where аналитика='+' order by ДО ";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DOO[i] = reader["ДО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return DOO;
        }

        public static string[] DOMotiv()
        {
            string[] DOO;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(ДО) from ДО where Мотивация='+'";
                    object value = cmd.ExecuteScalar();
                    DOO = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select ДО from ДО where Мотивация='+' order by ДО";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DOO[i] = reader["ДО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return DOO;
        }

        public static string[] UsersMotiv()
        {
            string[] DOO;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from Пользователи where Должность='+'";
                    object value = cmd.ExecuteScalar();
                    DOO = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select ФИО from Пользователи where Должность='+' order by ФИО";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DOO[i] = reader["ФИО"].ToString();
                            i++;
                        }
                    }
                }
            }
            return DOO;
        }


        public static string[] rules()
        {
            string[] rules;
            using (var conn = new NpgsqlConnection(data.path))
            {
                conn.Open();
                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = "select count(Код) from Права";
                    object value = cmd.ExecuteScalar();
                    rules = new string[Convert.ToInt32(value)];
                    cmd.CommandText = "select Значение from Права";
                    int i = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            rules[i] = reader["Значение"].ToString();
                            i++;
                        }
                    }
                }
            }
            return rules;
        }


        public static string[] doljnost()
        {
            String[] perms = data.userrules.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            return perms;
        }



        public static string managernum { get; set; } //Номерманагера
        public static string numberin { get; set; } // Номерзвонящего
        public static string adiingtype { get; set; } //Открытие добавления
        public static string idtask { get; set; } //id задагния
        public static string typeactiontask { get; set; } // новое или открытие задания
        public static string path { get; set; } // пусть к базе
        public static string ipfiles { get; set; } // ip для файлов
        public static int active { get; set; } //Активность
        public static int typeoflist { get; set; }// тип вывода списка
        public static string typesort { get; set; }
        public static string textsort { get; set; }
        public static int typequery  { get; set; } // тип выборки
        public static string query { get; set; } // выборка
        public static string cityzayavka  { get; set; } // для заявки 
        // public static int selecton = 0;
        public static string checkcode { get; set; }
        public static string checkaction { get; set; }
        public static string usercode { get; set; } // код пользователя
        public static string usercity { get; set; } // подразделениеъ
        public static string usercitydop { get; set; } // подразделение доп
        public static string usercityobzvon { get; set; } // подразделение для обзвона
        public static string userrules { get; set; } // права пользователя 
        public static string userFIO { get; set; } // ФИО пользователя
        public static string username { get; set; } // имя пользователя
        public static string userphone { get; set; } // тел пользователя
        public static string userpermission { get; set; } //роли юзера 
        public static string dogovorcode { get; set; } // код договора
        public static string dogovortablecode { get; set; }// код договора в таблице

        public static string attachfile { get; set; } // путь файла
        public static string namefile { get; set; } // имя файла
        public static string notiftext { get; set; } // текст уведомления
        public static string callaudio { get; set; }
        public static string selectzayavka { get; set; } // выболрка заявка
        public static string selectoplata { get; set; } // выболрка оплата
        public static string selectpretenzii { get; set; } // выболрка 76счет
        public static string selectdolg { get; set; } // выболрка долг
        public static string selectkons { get; set; } // выболрка конс
        public static string selectcalls { get; set; } // выболрка звонки
        public static string selectoldbase { get; set; } // выболрка старая база
        public static bool updatetable { get; set; } // обновлять ли таблицы
        public static string whatanaliz { get; set; } // какой анализ
        public static string analizpodr { get; set; } // Подразденление анализ
        public static string analizdateotch { get; set; } // Дата вывода анализа
        public static DateTime analizdatestart { get; set; }// Дата начала анализ
        public static DateTime analizdateend { get; set; }// Дата конец анализ
        public static string fioclient { get; set; } // фио клиента для фото
        public static string dogovorclient { get; set; } // договор для фото
        public static DateTime analizdateform { get; set; }// Дата формирования анализ
    }
}
