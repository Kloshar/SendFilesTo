// Copyright © 2024 Ilya Bashmakov. All rights reserved. Contacts: kloshar13@yahoo.com

using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Threading.Tasks;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Controls.Primitives;
using MimeKit;
using Microsoft.Win32;
using System.Security;
using System.Security.Permissions;

/*
Небольшое описание:
в этом файле задаются файлы для отправки и вызывается метод класса MailSender.cs,
который отправляет данные для авторизации, получает токен, извлекает код доступа
и отправляет переданное из этого метода сообщение

to do:
1. Отправка нескольких файлов. Например, если приложение уже запущено (пауза?), то добавлять файлы в список.
   Использование расширений проводника
2. Требуется запускать с правами администратора. В противном слуючае может быть ошибка прав доступа к реестру в local machine.
   Нужно добавить окошко в установку с предупреждением. И в справку.
 */

namespace SendFiles2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        public static MainWindow main; //переменная для закрытия окна из класса
        MimeMessage mm;
        MailSender ms;
        string fullName;
        string path;         
        string from = string.Empty;
        //string to = string.Empty;
        string[] to = new string[1];
        string[] arrayAttachments = new string[1];
        public Progress prog;

        public MainWindow()
        {
            /*
            //yahoo; //работает только через одноразовый пароль для приложения
            //mail; //работает только через одноразовый пароль для приложения
            */

            //первым делом получаем аргументы командной строки
            string[] args = Environment.GetCommandLineArgs();
            string fullArgs = string.Empty;
            foreach (string s in args)
            {
                fullArgs += s + Environment.NewLine;
            }

            //MessageBox.Show($"{args.Length} {fullArgs}");

            InitializeComponent();

            main = this;
            prog = new Progress();

            //получаем имя запускаемого файла с полным путём и выделяем путь
            fullName = Process.GetCurrentProcess().MainModule.FileName;
            path = fullName.Substring(0, fullName.LastIndexOf(@"\") + 1);

            //читаем адрес отправителя из файла fromArdess.txt, если он есть
            try
            {
                using (StreamReader sr = new StreamReader(path + @"fromArdess.txt"))
                {
                    from = sr.ReadLine();
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + "Возможно, требуется ввести адрес отправителя в главном окне программы...");
            }
            //отображаем адрес отправителя в окне
            textbox_from.Text = from;

            //если запущен экзешник, то будет передан один аргумент - путь к программе            
            if (args.Length == 1)
            {
                //загружаем адреса из файла в листбокс
                if (File.Exists(path + @"\adresses.txt"))
                {
                    FileInfo fi = new FileInfo(path + @"adresses.txt");
                    string[] adresses = File.ReadAllLines(path + @"adresses.txt");
                    addressesList.Items.Clear();
                    int count = adresses.Length;
                    for (int i = 0; i < count; i++) { addressesList.Items.Add(adresses[i]); }
                }

                //проверка интеграции в проводник
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"*\\shell", writable: true);

                //если ключ существует - ставим галочку
                if (key.OpenSubKey(@"SendFiles2") != null) { checkBox_integrate.IsChecked = true; }

                //задаём файлы и адрес для отправки (требуется только в случае отладки)
                arrayAttachments[0] = @"d:\log.txt"; //пока не придумал как передать сразу несколько аргументов

                //делим адрес пробелом
                //to =("kloshar13@yahoo.com Kloshar13@mail.ru").Split(' '); //вариант на два адреса
                to[0] = "kloshar13@yahoo.com";

                //button_send.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            }            
            //если запускается правой кнопкой на файле, то будет передано три или более аргументов: путь к программе, отправляемый файл, адрес или адреса получателя
            else if (args[1] == "-s")
            {
                //загружаем адреса из файла в листбокс
                if (File.Exists(path + @"\adresses.txt"))
                {
                    FileInfo fi = new FileInfo(path + @"adresses.txt");
                    string[] adresses = File.ReadAllLines(path + @"adresses.txt");
                    addressesList.Items.Clear();
                    int count = adresses.Length;
                    for (int i = 0; i < count; i++) { addressesList.Items.Add(adresses[i]); }
                }
                //MessageBox.Show($"Запуск с аргументом {args[1]}");
                main.Show();
                //проверка интеграции в проводник
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"*\\shell", writable: true);
                //если ключ существует - ставим галочку
                if (key.OpenSubKey(@"SendFiles2") != null) { checkBox_integrate.IsChecked = true; }
                Close();
            }            
            else
            {
                //скрываем главное окно
                main.Visibility = Visibility.Visible;
                //запускаем окно
                main.Show();
                //задаём файлы для проверки отправки
                arrayAttachments[0] = args[1]; //пока не придумал как передать сразу несколько аргументов
                                               //отправляем письмо
                                               //тут надо записать кому из аргументов в гллобальную переменную

                //получатели - все аргументы, кроме первых двух
                Array.Resize(ref to, args.Length - 2); //устанавливаем размер массива поолучателей
                
                Array.Copy(args, 2, to, 0, to.Length); //копируем часть аргументов в массив

                sendLetter(arrayAttachments);
            }
        }
        private async void sendLetter(string[] attachments)
        {
            //проверка адреса отправителя, если он не правильный, дальше можно не продолжать
            if (CheckAddress(from) != false)
            {
                //создаём экземпляр класса MailSender и передаём почтовый сервер в качестве аргумента
                //Console.WriteLine(from.Substring(from.LastIndexOf('@') + 1));
                switch (from.Substring(from.LastIndexOf('@') + 1))
                {
                    case "gmail.com":
                        ms = new MailSender("gmail.com");
                        break;
                    case "yandex.ru":
                        ms = new MailSender("yandex.ru");
                        break;
                    default:
                        Console.WriteLine("Почтовый провайдер не поддерживается! Используйте gmail.com или yandex.ru...");
                        MessageBox.Show("Почтовый провайдер не поддерживается! Используйте gmail.com или yandex.ru...");
                        break;
                }

                //создаём сообщение в виде MimeMessage
                mm = CreateMessage(attachments, from, to);

                //отправляем сообщение
                try
                {
                    //SendMailKit(mm, accessCode);
                    await ms.SendMailKit(mm); //начинаем отправку
                    await ms.saveToSentFolder(mm); //сохраняем в отправленные                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Адрес отправителя не соответствует шаблону..." + "\n" + "Введите адрес отправителя в главном окне программы...");
                Close();
            }
        }
        private void button_send_Click(object sender, RoutedEventArgs e)
        {
            //отправляем письмо
            sendLetter(arrayAttachments);
        }
        private MimeMessage CreateMessage(string[] attachments, string from, string[] to)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(from, from));

            //здесь нужно сделать цикл
            foreach(string r in to)
            {
                message.To.Add(new MailboxAddress(r, r));
            }
            message.Subject = "Файл отправлен от " + from;
            BodyBuilder builder = new BodyBuilder();
            builder.TextBody = @"Это сообщение отправлено через SendFiles2!";
            foreach (string file in attachments)
            {
                builder.Attachments.Add(file);
            }
            message.Body = builder.ToMessageBody();

            return message;
        }
        private bool CheckAddress(string address)
        {
            //проверка адреса на соответствие шаблону
            Regex regex = new Regex(@"\w+\@\w+\.\w+"); //\w - любой символ, + - ниличие одного или более этих символов, точка и собака экранируются слешем
            Match match = regex.Match(address);
            return match.Success;
        }
        private void addressesList_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Delete)
            {
                int selectedIndex = addressesList.SelectedIndex;
                addressesList.Items.RemoveAt(selectedIndex);
                //using (StreamWriter sw = new StreamWriter(@"adresses.txt"))
                //{
                //    sw.WriteLine();
                //}

                int count = addressesList.Items.Count;
                string[] addreses = new string[count];
                for (int i = 0; i < count; i++) { addreses[i] = addressesList.Items[i].ToString(); }
                File.WriteAllLines(path + @"\adresses.txt", addreses);

                //нужно заново поставить галочку
                checkBox_integrate.IsChecked = false;
                checkBox_integrate.IsChecked = true;
            }
        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {            
            //ловим исключение запрещения доступа к реестру
            try
            {
                //Console.WriteLine("Начинается добавление в реестр...");
                RegistryKey key, key1;

                //!!! возможно права создают проблемы с запуском на некоторых компьютерах

                //Представляет пользователя Windows
                //WindowsIdentity identify = WindowsIdentity.GetCurrent();
                //Обеспечивает безопасность управления доступом Windows для раздела реестра
                //RegistrySecurity regSecurity = new RegistrySecurity();
                //Представляет набор прав доступа, разрешенных или запрещенных пользователю или группе
                //RegistryAccessRule accessRule = new RegistryAccessRule(identify.User, RegistryRights.FullControl, AccessControlType.Allow);
                //Удаляет все правила управления доступом с тем же именем пользователя и значением свойства AccessControlType ("разрешить" или "запретить"), 
                //что и у указанного правила, после чего добавляет указанное правило
                //regSecurity.SetAccessRule(accessRule);

                key = Registry.ClassesRoot.OpenSubKey(@"*\\shell", writable: true);

                //открываем ключ для регистрации ярлыков поддиректорий. Небольная фигня с 64-разрядной виндой
                key1 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\", writable: true);

                //добавляем в реестр информацию для контестного меню
                //создаём подключ реестра в ClassesRoot\*\Shell\SendFiles2
                key = key.CreateSubKey(@"SendFiles2", RegistryKeyPermissionCheck.ReadWriteSubTree/*, regSecurity*/);
                //создаём строковый параметр с названием пунктов
                key.SetValue("MUIVerb", "SendFiles2..."); //название
                key.SetValue("Icon", Process.GetCurrentProcess().MainModule.FileName); //значок
                key.SetValue("SubCommands", ""); //действие

                //переменная для списка поддиректорий каскадного меню в *\shell
                string sub_commands = string.Empty;

                for (int i = 0; i < addressesList.Items.Count; i++)
                {
                    //если строка с командами пуста, то добавляем первый элемент
                    if (sub_commands == "")
                    {
                        sub_commands = addressesList.Items[i].ToString();
                    }
                    //если уже что-то содержит, то добавляем точку с запятой и элемент
                    else
                    {
                        sub_commands += ";" + addressesList.Items[i];
                    }
                    //записываем команды в подключ SubCommands
                    key.SetValue("SubCommands", sub_commands);
                    //создаём ключи для регистрации ярлыков поддиректорий
                    key1 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\", writable: true);
                    key1 = key1.CreateSubKey(addressesList.Items[i].ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree/*, regSecurity*/);
                    key1.SetValue("", addressesList.Items[i]);
                    //создаём ключ command для запуска
                    key1 = key1.CreateSubKey(@"Command");
                    //путь к программе
                    string exe_path = path + @"SendFiles2.exe ""%1""" + " " + addressesList.Items[i];
                    //записываем данные в ключ
                    key1.SetValue("", exe_path);
                }
                key.Close();
                key1.Close();
            }
            catch (SecurityException ex)            
            {
                MessageBox.Show(ex.Message + "\nДля корректной работы требуется запускать SendFiles2.exe от имени Администратора!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            RegistryKey key, key1;
            key = Registry.ClassesRoot.OpenSubKey(@"*\\shell", writable: true);
            key1 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\", writable: true);

            //удаляем лишние разделы из реестра, если галочка снимается
            //для начала прочитаем ключи из части реестра hkcr
            //key = key.OpenSubKey(@"SendFiles2", writable: true);
            if (key.OpenSubKey(@"SendFiles2") != null && key.OpenSubKey(@"SendFiles2", writable: true).GetValue("SubCommands") != null)
            {
                string sub_commands = key.OpenSubKey(@"SendFiles2", writable: true).GetValue("SubCommands").ToString();
                string[] sub_commands_array;
                sub_commands_array = sub_commands.Split(new char[] { ';' });
                //далее удаляем их из части реестра hklm
                for (int i = 0; i < sub_commands_array.Length; i++)
                {
                    try
                    {
                        key1.DeleteSubKeyTree(sub_commands_array[i]);
                    }
                    catch (ArgumentException ex)
                    {
                        Console.Write(ex);
                    }
                }
                //затем стираем и раздел в части hkcr
                key.DeleteSubKey(@"SendFiles2");
            }
        }
        private void button_add_Click(object sender, RoutedEventArgs e)
        {
            //добавляет адрес из строки в список доступных для отправки адресов
            if (addressBox.Text != "")
            {
                addressesList.Items.Add(addressBox.Text);
                addressBox.Text = "";

                int count = addressesList.Items.Count;
                string[] addreses = new string[count];
                for (int i = 0; i < count; i++) { addreses[i] = addressesList.Items[i].ToString(); }
                File.WriteAllLines(path + @"\adresses.txt", addreses);

                //нужно заново поставить галочку
                checkBox_integrate.IsChecked = false;
                checkBox_integrate.IsChecked = true;

                addressBox.Focus();
            }
        }        
        private void FirstWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string addr = textbox_from.Text;
            //если поле адреса заполнено, то сохраняем его в файл
            if(addr != "" && CheckAddress(addr))
            {
                using (StreamWriter sw = new StreamWriter(path + @"fromArdess.txt", false))
                {
                    sw.WriteLine(textbox_from.Text);
                }
            }
            Application.Current.Shutdown();
        }
    }
}