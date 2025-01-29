// Copyright © 2024 Ilya Bashmakov. All rights reserved. Contacts: kloshar13@yahoo.com

using System;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Util;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MailKit.Net.Imap;
using MailKit;
/*
класс предназначен для получения токена,
извлечения кода из токена
и отправки письма
Использование: 
создать письмо MimeMessage mm,
создать экземпляр класса, в конструкторе передать имя почтового сервера
вызвать метод SendMailKit и передать ему письмо SendMailKit(mm)
*/
namespace SendFiles2
{
    class MailSender
    {
        static readonly Guid SID_SWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
        string provider = "";
        string client_id = "";
        string client_secret = "";
        string access_reference = "";
        string code = "";
        string token = "";
        Window w;

        string fromAddress = string.Empty;
        string smtp = "";
        string imap = "";
        int smtpPort = 587;
        int imapPort = 993;
        string accessCode = string.Empty;

        //string login = "";
        //string password = "";

        string fullName;
        string path;
        string credPath;

        string from = string.Empty;

        MainWindow mw; // была нужна для закрытия главного окна
        Progress pro;
        UserCredential credential;

        public MailSender(string mailProvider) 
        { 
            provider = mailProvider;

            //prog = new Progress();
            mw = MainWindow.main;
            pro = mw.prog;

        }
        public async Task SendMailKit(MimeMessage mm)
        {
            //для авторизации нужно в настройках почты установить галочки: Почтовые программы -> разрешить доступ с помощью почтовых клиентов ->
            //Способ авторизациипо IMAP Пароли приложений и OAuth-токены
            //https://mail.yandex.ru/?uid=132231255#setup/client

            //MessageBox.Show("!");

            pro.Show();

            pro.labelProg.Content = "Начата отправка";
            pro.progressBar.Value = 0;

            //адрес отправителя  берём из сформированного письма
            fromAddress = ((MailboxAddress)mm.From[0]).Address;

            pro.labelProg.Content = $"Получен адрес отправителя: {fromAddress}";
            pro.progressBar.Value += 10;

            //Авторизация пользовател. Сохраяется переменная token
            Authorization();

            //извлекаем код доступа из переменной token
            accessCode = ExtractAccessCode();

            //подставляем адрес и код с объект авторизации
            SaslMechanismOAuth2 oauth2 = new SaslMechanismOAuth2(fromAddress, accessCode);

            SmtpClient smtpClient = new SmtpClient(); //smtp client
            await smtpClient.ConnectAsync(smtp, smtpPort, SecureSocketOptions.StartTls); //smtp client

            try
            {
                pro.labelProg.Content = $"Аутентификация почтового клиента через OAuth2";
                pro.progressBar.Value += 20;

                await smtpClient.AuthenticateAsync(oauth2); //авторизация по oauth //smtp client                

                //await smtpClient.AuthenticateAsync(SenderAdress, "7p74cJUjbE3tLsXKv0Qe"); //авторизация по логину и паролю                

                pro.labelProg.Content = $"Отправка письма";
                pro.progressBar.Value += 20;

                //отправка сообщения
                await smtpClient.SendAsync(mm); //smtp client

                //отключение
                await smtpClient.DisconnectAsync(true); //smtp client

                //получаем имя прикреплённого файла для того, чтобы написать его в MessageBox после отправки
                string fileName = string.Empty;
                foreach (var attachment in mm.Attachments)
                {
                    if(attachment is MessagePart)
                    {
                        fileName = attachment.ContentDisposition?.FileName;
                        Console.WriteLine(fileName);

                        MessagePart rfc822 = (MessagePart)attachment;

                        if (string.IsNullOrEmpty(fileName)) { fileName = "attachment-message.eml"; }
                        Console.WriteLine(fileName);
                    }
                    else
                    {
                        MimePart part = (MimePart)attachment;
                        fileName = part.FileName;
                        Console.WriteLine($"Часть сообщения: {fileName}");
                    }
                }

                //перезапись обновлённого токена
                rewriteToken(token);

                pro.labelProg.Content = $"Сообщение отправлено!";
                pro.progressBar.Value = 100;

                MessageBox.Show($"Сообщение отправлено! Файл: {fileName}, получатель: {mm.To}", "Результат отправки", MessageBoxButton.OK);               

            }
            catch (AuthenticationException Authentication_ex)
            {
                Console.WriteLine(Authentication_ex.Message);
                StringBuilder errorStr = new StringBuilder(Authentication_ex.Message);
                errorStr.AppendLine($"Адрес отправителя: {fromAddress}");
                errorStr.AppendLine($"Код доступа: {accessCode}");

                MessageBox.Show(errorStr.ToString(), Authentication_ex.ToString(), MessageBoxButton.OK, MessageBoxImage.Stop);

                // функция удаления токена гугл повторяется в методе аутентификации (сложно сказать где её правильно оставить)
                if (provider == "gmail.com")
                {
                    //по идее, если будет ошибка авторизации, то нужно стереть папку с токеном
                    //функция не проверена, так сложно сделать такую ошибку специально
                    DirectoryInfo tokenDir = new DirectoryInfo(path + @"token.json");
                    if (Directory.Exists(tokenDir.FullName))
                    {
                        tokenDir.Delete(true);
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //MainWindow.main.Close();

            Console.WriteLine("Метод отправки завершён!");

            /*
             сохранение в отправленные https://github.com/jstedfast/MailKit/blob/master/FAQ.md#SpecifiedPickupDirectory
            */
        }
        public async Task saveToSentFolder(MimeMessage mm)
        {
            Console.WriteLine("В методе saveToSentFolder");

            pro.labelProg.Content = $"Сохранение в Отправленные";
            pro.progressBar.Value += 10;

            //адрес отправителя  берём из сформированного письма
            fromAddress = ((MailboxAddress)mm.From[0]).Address;

            //Авторизация пользовател. Сохраяется переменная token
            Authorization();

            //извлекаем код доступа из переменной token
            accessCode = ExtractAccessCode();

            //подставляем адрес и код с объект авторизации
            SaslMechanismOAuth2 oauth2 = new SaslMechanismOAuth2(fromAddress, accessCode);

            //создаём imap-клиент
            ImapClient imapClient = new ImapClient();

            imapClient.CheckCertificateRevocation = false;

            //подключаемся
            await imapClient.ConnectAsync(imap, imapPort, SecureSocketOptions.Auto);

            //авторизуемся
            await imapClient.AuthenticateAsync(oauth2);

            //выводим папки
            IList<IMailFolder> folders = await imapClient.GetFoldersAsync(imapClient.PersonalNamespaces.First());

            //Console.WriteLine($"Обнаружено папок в folders: {folders.Count}");
            //foreach (IMailFolder folder in folders)
            //{
            //    Console.WriteLine(folder.Name + " - " + folder.Id + folder.Attributes + " ");
            //}

            IMailFolder sent = null;
            
            //это не работает в яндекс, так как возвращает false
            if (imapClient.Capabilities.HasFlag(ImapCapabilities.SpecialUse))
            {                
                sent = imapClient.GetFolder(SpecialFolder.Sent);
            }            

            if (sent == null)
            {
                IMailFolder personal = imapClient.GetFolder(imapClient.PersonalNamespaces.First()); //получаем персональные папки
                sent = personal.GetSubfolder("Отправленные"); //получаем папку отправленные
                UniqueId? uid = await sent.AppendAsync(mm, MessageFlags.Seen); //добавляем к папке письмо
                Console.WriteLine(uid);
            }
            await imapClient.DisconnectAsync(true);

            pro.labelProg.Content = $"Сообщение сохранено в Отправленные!";

            MainWindow.main.Close();
        }
        public string ExtractAccessCode()
        {            
            if (provider == "yandex.ru")
            {
                //из строки с токеном нужно извлечь только код доступа, создаём объект (класс) tok с нужным полем access_token
                var tok = JsonConvert.DeserializeAnonymousType(token, new { access_token = "" });
                accessCode = tok.access_token;
            }
            if (provider == "gmail.com")
            {
                //сохраняем данные в json объект
                JObject gtoken = (JObject)JsonConvert.DeserializeObject(token);
                
                //извлекаем код доступа из токена
                accessCode = (string)gtoken["access_token"];
            }
            if (provider == "yahoo.com")
            {
                var tok = JsonConvert.DeserializeAnonymousType(token, new { access_token = "" });
                accessCode = tok.access_token;
            }
            if (provider == "mail.ru")
            {
                //из строки с токеном нужно извлечь только код доступа, создаём объект (класс) tok с нужным полем access_token
                int i = token.IndexOf("&access_token=");
                string str = token.Substring(i + 14);
                accessCode = str.Substring(0, str.IndexOf('&'));
            }

            pro.labelProg.Content = $"Код доступа извлечён";
            pro.progressBar.Value += 10;

            //Console.WriteLine($"Код доступа: { accessCode }");
            return accessCode;
        }
        public string Authorization()
        {
            //результат работы этого метода - присваивание переменной token
            //в зависимости от почтового сервера выбираем метод получения токена

            //получаем имя запускаемого файла с полным путём и выделяем путь
            fullName = Process.GetCurrentProcess().MainModule.FileName;
            path = fullName.Substring(0, fullName.LastIndexOf(@"\") + 1);

            //Console.WriteLine(path);

            switch (provider)
            {
                case "yandex.ru":
                    {
                        //присваиваем значения переменным id и secret
                        client_id = "a66a4ac773854c0182426823ded214cd";
                        client_secret = "b0ce7014f1534f4d88fbd017db2acb72";

                        //ссылка для перенаправления пользователя на авторизацию
                        access_reference = @"https://oauth.yandex.ru/authorize?" + "response_type=code" + "&client_id=" + client_id;
 
                        //путь для записи токена
                        credPath = path + @"yandex_token.json";

                        //если файл с токеном доступа существует, то открываем,
                        if (File.Exists(credPath))
                        {
                            using (StreamReader sr = new StreamReader(credPath))
                            {
                                token = sr.ReadToEnd();
                            }
                        }
                        //если нет, то получаем доступ и записываем в файл
                        else
                        {
                            Thread newThread = new Thread(this.newWindow);
                            newThread.SetApartmentState(ApartmentState.STA);
                            newThread.Start();
                        }
                        //приостанавливаем этот поток, чтобы дождаться записи значения переменной token (в методе b_loaded)
                        for (int i = 0; token == "" & i <= 300; i += 1)
                        {
                            Thread.Sleep(1000);
                            Console.WriteLine("Ожидание... Прошло {0} сек. Токен равен {1}!", i, token);

                            //отключил подтверждение ожидания, чтобы не мешало вводить данные в форму авторизации
                            //if (i >= 5)
                            //{
                            //    //нужно дать выбор продолжать ожидание или нет...
                            //    if (token == "") w.Dispatcher.Invoke(
                            //        new Action(
                            //            () =>
                            //            {
                            //                if (MessageBox.Show(w, "Прошло 5 секунд! Если хотите продолжить ожидание - нажмите YES", "Ошибка", MessageBoxButton.YesNo, MessageBoxImage.Error) == MessageBoxResult.Yes)
                            //                {
                            //                    i = 0;
                            //                }
                            //            }));
                            //}
                        }
                        //Console.WriteLine($"Токен получен: { token }");

                        //заодно устанавливаем переменные
                        smtp = "smtp.yandex.ru";
                        imap = "imap.yandex.ru";  
                    }
                    break;
                case "gmail.com":
                    {
                        //объявляем переменные для хранения названия приложения                
                        string ApplicationName = "SendFiles2";

                        //присваиваем значения переменным id и secret
                        client_id = "587932262187-b1k12m50784nj7sl49sk9pc30an4a81p.apps.googleusercontent.com";
                        client_secret = "PO-6Wyjvf3ry7AII6nz3m3Gb";

                        // Переменная, описывающая id и большой секрет
                        ClientSecrets secr = new ClientSecrets { ClientId = client_id, ClientSecret = client_secret };

                        //права приложения
                        string[] scopes = { GmailService.Scope.MailGoogleCom };

                        //путь для записи токена
                        credPath = path + @"token.json";

                        //если токен в папке "token.json"
                        if (Directory.Exists(credPath) && Directory.GetFiles(credPath).Length > 0)
                        {
                            //путь к единственному файлу в папке                        
                            string tokenPath = Directory.GetFiles(credPath).First();

                            //читаем токен из папки
                            using (StreamReader sr = new StreamReader(tokenPath))
                            {
                                token = sr.ReadToEnd();
                            }
                            
                            //сохраняем данные в json объект
                            JObject gtoken = (JObject)JsonConvert.DeserializeObject(token);

                            //проверка времени истечения токена
                            if (isActual(tokenPath))
                            {
                                Console.WriteLine("Токен актуален!"); //ничего не делаем, переменная token уже присвоена
                            }
                            else
                            {
                                Console.WriteLine("Токен устарел!");
                                //записываем данные из устаревшего токена в tokenResponse
                                TokenResponse refreshToken = new TokenResponse
                                {
                                    AccessToken = (string)gtoken["access_token"],
                                    //ExpiresInSeconds = (long)gtoken["expires_in"],
                                    Scope = (string)gtoken["https://mail.google.com/"],                                    
                                    RefreshToken = (string)gtoken["refresh_token"],
                                    //IssuedUtc = (DateTime)gtoken["IssuedUtc"],
                                    //TokenType = (string)gtoken["token_type"]
                                };

                                //создаём безопасный поток авторизации с секретами и доступом
                                IAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                                {
                                    ClientSecrets = secr,
                                    Scopes = scopes
                                }
                                );

                                //создаём учётные данные пользователя из потока, передавая токен для обновления
                                credential = new UserCredential(flow, fromAddress, refreshToken);

                                //обновляем токен через учётные данные функцией RefreshTokenAsync
                                try
                                {
                                    bool success = credential.RefreshTokenAsync(CancellationToken.None).Result;  //если доступ отозван, то ошибка
                                                                                                                 //создаём gmail api сервис для авторизации и для отправки письма
                                    GmailService service = new GmailService(new BaseClientService.Initializer()
                                    {
                                        HttpClientInitializer = credential,
                                        ApplicationName = ApplicationName
                                    });

                                    //записываем обновлённый токен в переменную
                                    token = JsonConvert.SerializeObject(credential.Token);
                                }
                                catch (AggregateException ex)
                                {
                                    //выводить оконшко ошибки смысла нет, только лишний клик будет
                                    //foreach(var e in ex.InnerExceptions) MessageBox.Show("Ошибка обновления токена! После автоматического удаления токена требуется повторный вход в аккаунт \n" + e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Stop);                                    

                                    Console.WriteLine(ex.Message);

                                    //по идее, если будет ошибка авторизации, то нужно стереть папку с токеном
                                    DirectoryInfo tokenDir = new DirectoryInfo(path + @"token.json");
                                    if (Directory.Exists(tokenDir.FullName))
                                    {
                                        tokenDir.Delete(true);
                                    }

                                    //учётные данные пользователя, получаем токен и сохраняем на диск
                                    googleAutorize(secr, scopes, credPath, ApplicationName, credential);
                                    //Environment.Exit(0);
                                }
                            }
                        }
                        //если токена нет в папке "token.json", то получаем его через авторизацию пользователя
                        else
                        {
                            //учётные данные пользователя, получаем токен и сохраняем на диск
                            googleAutorize(secr, scopes, credPath, ApplicationName, credential);
                        }

                        //    Вариант отправки через сервис gMail
                        //    SendMsgGmail(service, credential);
                        //    string plaintext = "To: kloshar13@yahoo.com\r\n" +
                        //        "Subject: test\r\n" +
                        //        "Content-Type: text/html; charset=us-ascii\r\n\r\n" +
                        //        "<h1>Это сообщение отпралено через Gmail API<h1>";
                        //    var gMessage = new Google.Apis.Gmail.v1.Data.Message();
                        //    gMessage.Raw = Base64UrlEncode(plaintext.ToString());
                        //    service.Users.Messages.Send(gMessage, credential.UserId).Execute();

                        //заодно устанавливаем переменные
                        smtp = "smtp.gmail.com";
                        imap = "imap.gmail.com";
                    }
                    break;
                case "yahoo.com":
                    {
                        /*
                         Не удалось реализовать отправку. Требуется дать приложению права на запись. Yahoo отказал в запросе.
                         То есть токен получить можно, но письмо по нему не будет отправлено.
                         */
                        //присваиваем значения переменным id и secret
                        client_id = "dj0yJmk9RFJ0ZUhqeWVhTTBEJmQ9WVdrOVQxUkxWVXRuZVdFbWNHbzlNQT09JnM9Y29uc3VtZXJzZWNyZXQmc3Y9MCZ4PTVl";
                        client_secret = "7f77e19b9ae6a46bd445517f6db58344ec66930e";

                        //ссылка для перенаправления пользователя на авторизацию
                        access_reference = @"https://api.login.yahoo.com/oauth2/request_auth?" + "client_id=" + client_id + "&redirect_uri=oob" + "&response_type=code" + "&language=en-us";
                        //если файл с токеном доступа существует, то открываем, 
                        if (File.Exists("yahoo_token.json"))
                        {
                            using (StreamReader sr = new StreamReader("yahoo_token.json"))
                            {
                                token = sr.ReadToEnd();
                            }
                        }
                        //если нет, то получаем доступ и записываем в файл
                        else
                        {
                            Thread newThread = new Thread(this.newWindow);
                            newThread.SetApartmentState(ApartmentState.STA);
                            newThread.Start();
                        }
                        //приостанавливаем этот поток, чтобы дождаться записи значения переменной token
                        for (int i = 0; token == "" & i <= 5; i += 1)
                        {
                            Thread.Sleep(1000);
                            Console.WriteLine("Ожидание... Прошло {0} сек. Токен равен {1}!", i, token);
                            if (i == 5)
                            {
                                //нужно дать выбор продолжать ожидание или нет...
                                if (token == "") w.Dispatcher.Invoke(
                                    new Action(
                                        () =>
                                        {
                                            if (MessageBox.Show(w, "Прошло 5 секунд! Если хотите продолжить ожидание - нажмите YES", "Ошибка", MessageBoxButton.YesNo, MessageBoxImage.Error) == MessageBoxResult.Yes)
                                            {
                                                i = 0;
                                            }
                                        }));
                            }
                        }

                        //заодно устанавливаем переменные
                        smtp = "smtp.mail.yahoo.com";

                        //login = "kloshar13";
                        //password = "gmgawrcbgonxiarg"; //одноразовый пароль
                    }
                    break;
                case "mail.ru":
                    {
                        /*
                        Не удалось реализовать отправку. Подозреваю, что ситуация такая же как в yahoo, так как нет документации по отправке с OAuth токеном.
                        То есть токен получить можно, но письмо по нему не будет отправлено.
                        */
                        //присваиваем значения переменным id и secret
                        client_id = "784997";
                        client_secret = "2ac3e6da417a9b53a135e42ced085fe3";

                        //ссылка для перенаправления пользователя на авторизацию
                        access_reference = @"https://connect.mail.ru/oauth/authorize?" + "&client_id=" + client_id + "&response_type=token";

                        //если файл с токеном доступа существует, то открываем, 
                        if (File.Exists("mailru_token.json"))
                        {
                            using (StreamReader sr = new StreamReader("mailru_token.json"))
                            {
                                token = sr.ReadToEnd();
                            }
                        }
                        //если нет, то получаем доступ и записываем в файл
                        else
                        {
                            Thread newThread = new Thread(this.newWindow);
                            newThread.SetApartmentState(ApartmentState.STA);
                            newThread.Start();
                        }
                        //приостанавливаем этот поток, чтобы дождаться записи значения переменной token
                        for (int i = 0; token == "" & i <= 20; i += 1)
                        {
                            Thread.Sleep(1000);
                            Console.WriteLine("Ожидание... Прошло {0} сек. Токен равен {1}!", i, token);
                            if (i == 20)
                            {
                                //нужно дать выбор продолжать ожидание или нет...
                                if (token == "") w.Dispatcher.Invoke(
                                    new Action(
                                        () =>
                                        {
                                            if (MessageBox.Show(w, "Прошло 5 секунд! Если хотите продолжить ожидание - нажмите YES", "Ошибка", MessageBoxButton.YesNo, MessageBoxImage.Error) == MessageBoxResult.Yes)
                                            {
                                                i = 0;
                                            }
                                        }));
                            }
                        }

                        smtp = "smtp.mail.ru";

                    }
                    break;
                default:
                    Console.WriteLine("Не выбран провадер!");
                    break;
            }

            pro.labelProg.Content = $"Токен получен";
            pro.progressBar.Value += 20;

            Console.WriteLine(token);
            return token;
        }
        //авторизация в гугл и получение учётных данных пользователя
        private void googleAutorize(ClientSecrets secr, string[] scopes, string credPath, string ApplicationName, UserCredential credential)
        {
            //учётные данные пользователя, получаем токен и сохраняем на диск
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                secr, scopes, fromAddress, CancellationToken.None, new FileDataStore(credPath, true)
                ).Result;

            pro.labelProg.Content = $"Авторизация пользователя. Токен получен";
            pro.progressBar.Value += 10;

            //создаём gmail api сервис для авторизации и для отправки письма
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
            token = JsonConvert.SerializeObject(credential.Token);
        }
        //перезапись обновлённого токена
        private void rewriteToken(string tkn)
        {
            if (provider == "yandex.ru")
            {
                //не требуется, так как токен не устарел за почти год
            }
            if (provider == "gmail.com")
            {
                if (!Directory.Exists(credPath)) Directory.CreateDirectory(credPath);

                using (StreamWriter sw = new StreamWriter(credPath + @"\Google.Apis.Auth.OAuth2.Responses.TokenResponse-" + fromAddress))
                {
                    sw.WriteLine(token);
                }
            }
        }
        //функция определяет просрочен ли токен
        bool isActual(string tokenPath)
        {
            DateTime dt = new DateTime(); //дата истечения токена

            string data = string.Empty; //данные токена

            //читаем токен из папки
            using (StreamReader sr = new StreamReader(tokenPath))
            {
                data = sr.ReadToEnd();
            }

            //сохраняем данные в json объект
            JObject gtoken = (JObject)JsonConvert.DeserializeObject(data);

            //извлекаем дату выдачи токена
            dt = (DateTime)gtoken["Issued"];

            //прибавляем время жизни токена
            dt = dt.AddSeconds(Convert.ToDouble(gtoken["expires_in"]));

            //сравниваем с текущей датой
            return DateTime.Now < dt;
        }
        void newWindow()
        {
            w = new Window();
            w.Height = 600;
            w.Width = 500;
            w.Title = "Token receiving";
            w.Closed += (object sender, EventArgs e) =>
            {
                Console.WriteLine("Окно закрыто. " + token);
            };

            WebBrowser b = new WebBrowser();
            b.Height = 500;
            b.Width = 500;
            b.LoadCompleted += b_loaded;
            b.Navigated += b_navigated;

            StackPanel s = new StackPanel();
            s.Children.Add(b);

            w.Content = s;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(access_reference);
            b.Navigate(((HttpWebResponse)request.GetResponse()).ResponseUri);

            w.ShowDialog();
        }
        void b_loaded(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            if (provider == "yandex.ru")
            {
                //записываем в переменную адрес перехода в браузере
                string str = e.Uri.ToString();
                //ищем код из ответа в адресе
                int ind = str.LastIndexOf("?code=");
                //если есть код, то записываем
                if (ind > -1)
                {
                    code = str.Substring(ind + 6);
                    //если код записан, то формируем и отправляем запрос
                    if (code != "")
                    {
                        //формируем строку запроса на яндекс
                        string data = string.Concat("grant_type=authorization_code", "&code=" + code, "&client_id=" + client_id, "&client_secret=" + client_secret);
                        //кодируем строку в последовательность байтов
                        byte[] databytes = Encoding.UTF8.GetBytes(data);
                        //создаём запрос и указывем тип отправляемых данных, метод запроса, тип метода распакавки и длину данных запроса
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://oauth.yandex.ru/token");
                        request.ContentType = "application/x-www-form-urlencoded"; //тип содержимого
                        request.Method = "POST"; //метод отправки
                        request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                        request.ContentLength = databytes.Length;
                        //записываем данные в поток запроса
                        using (Stream dataStream = request.GetRequestStream())
                        {
                            dataStream.Write(databytes, 0, databytes.Length);
                        }
                        //выполняем запрос
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        //читаем ответ из потока
                        using (Stream stream = response.GetResponseStream())
                        {
                            using (StreamReader reader = new StreamReader(stream))
                            {
                                //записываем токен в переменную
                                token = reader.ReadToEnd();
                            }
                        }
                        response.Close();
                        //закрываем окно
                        w.Close();
                    }
                    using (StreamWriter sw = new StreamWriter(credPath))
                    {
                        Console.WriteLine("Начинаем запись в файл...");
                        sw.Write(token);
                    }
                }
            }
            if (provider == "yahoo.com") 
            {
                //записываем в переменную адрес перехода в браузере
                string str = e.Uri.ToString();
                //ищем код из ответа в адресе
                int ind = str.LastIndexOf("&code=");
                Console.WriteLine(ind);
                //если есть код, то записываем
                if (ind > -1)
                {
                    code = str.Substring(ind + 6);
                    code = code.Substring(0, code.IndexOf("&"));
                    //если код записан, то формируем и отправляем запрос
                    if (code != "")
                    {
                        //формируем строку запроса на яху
                        string data = string.Concat("grant_type=authorization_code", "&redirect_uri=oob", "&code=" + code);
                        //кодируем строку в последовательность байтов
                        byte[] databytes = Encoding.UTF8.GetBytes(data);
                        //создаём запрос и указывем тип отправляемых данных, метод запроса, тип метода распакавки и длину данных запроса
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.login.yahoo.com/oauth2/get_token");
                        request.ContentType = "application/x-www-form-urlencoded"; //тип содержимого
                        request.Method = "POST"; //метод отправки
                        request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                        request.ContentLength = databytes.Length;
                        request.Headers.Add("Authorization", "Basic " + Base64UrlEncode(client_id + ":" + client_secret));
                        
                        //записываем данные в поток запроса
                        using (Stream dataStream = request.GetRequestStream())
                        {
                            dataStream.Write(databytes, 0, databytes.Length);
                        }
                        //выполняем запрос
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        //читаем ответ из потока
                        using (Stream stream = response.GetResponseStream())
                        {
                            using (StreamReader reader = new StreamReader(stream))
                            {
                                //записываем токен в переменную
                                token = reader.ReadToEnd();
                            }
                        }
                        response.Close();
                        //закрываем окно
                        w.Close();
                    }
                    using (StreamWriter sw = new StreamWriter("yahoo_token.json"))
                    {
                        Console.WriteLine("Начинаем запись в файл...");
                        sw.Write(token);
                    }
                }
            }
            if (provider == "mail.ru")
            {
                //записываем в переменную адрес перехода в браузере
                string str = e.Uri.ToString();
                //ищем код из ответа в адресе
                int ind = str.LastIndexOf("#");
      
                Console.WriteLine(token);
                //если есть код, то записываем
                if (ind > -1)
                {
                    //на майл ру сразу выдаётся токен, ничего обменивать не надо
                    token = str.Substring(ind + 6);
                    w.Close();
                    using (StreamWriter sw = new StreamWriter("mailru_token.json"))
                    {
                        Console.WriteLine("Начинаем запись в файл...");
                        sw.Write(token);
                    }
                }
            }
            //этот код подавляет открытие нового окна IE
            IServiceProvider serviceProvider = (IServiceProvider)((WebBrowser)sender).Document;
            Guid serviceGuid = SID_SWebBrowserApp;
            Guid iid = typeof(SHDocVw.IWebBrowser2).GUID;
            SHDocVw.IWebBrowser2 myWebBrowser2 = (SHDocVw.IWebBrowser2)serviceProvider.QueryService(ref serviceGuid, ref iid);
            SHDocVw.DWebBrowserEvents_Event wbEvents = (SHDocVw.DWebBrowserEvents_Event)myWebBrowser2;
            wbEvents.NewWindow += new SHDocVw.DWebBrowserEvents_NewWindowEventHandler(OnWebBrowserNewWindow);
            void OnWebBrowserNewWindow(string URL, int Flags, string TargetFrameName, ref object PostData, string Headers, ref bool Processed)
            {
                Processed = true;
                ((WebBrowser)sender).Navigate(URL);
            }
        }
        void b_navigated(object sender, EventArgs e)
        {
            //этот код подавляет выдачу сообщении об ошибках сценариев
            FieldInfo fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fiComWebBrowser == null) return;
            object objComWebBrowser = fiComWebBrowser.GetValue((WebBrowser)sender);
            if (objComWebBrowser == null) return;
            objComWebBrowser.GetType().InvokeMember("Silent", BindingFlags.SetProperty, null, objComWebBrowser, new object[] { true });
        }
        //эта хрень для подавления открытия нового окна IE
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
        internal interface IServiceProvider
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object QueryService(ref Guid guidService, ref Guid riid);
        }
        private static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            // Special "url-safe" base64 encode.
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }
        public void showFolders(GmailService service) //код получения списка папок
        {
            //создаём запрос
            UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List("me");

            //создаём список папок и заполняем его выполняя запрос
            IList<Google.Apis.Gmail.v1.Data.Label> labels = request.Execute().Labels;

            //выводим список папок почтового ящика
            if (labels != null && labels.Count > 0)
            {
                foreach (var labelItem in labels)
                {
                    Console.WriteLine(labelItem.Name);
                    //textbox.Text += labelItem.Name + Environment.NewLine;
                }
            }
            else
            {
                Console.WriteLine("No labels found.");
            }
        }
       
    }
}
