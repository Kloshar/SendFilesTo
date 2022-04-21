using System;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Text;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.Apis.Util;
using Newtonsoft.Json;


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

        public MailSender(string mailProvider) { provider = mailProvider; }
        public string Authorization()
        {
            //в зависимости от почтового сервера выбираем метод получения токена
            if(provider == "yandex.ru")
            {
                //присваиваем значения переменным id и secret
                client_id = "d85a4ac773854c0182426823ded235lk";
                client_secret = "dfr44ac773854c0182426823ded228hy";

                //ссылка для перенаправления пользователя на авторизацию
                access_reference = @"https://oauth.yandex.ru/authorize?" + "response_type=code" + "&client_id=" + client_id;

                //если файл с токеном доступа существует, то открываем, 
                if (File.Exists("yandex_token.json"))
                {
                    using (StreamReader sr = new StreamReader("yandex_token.json"))
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
            }
            return token;
        }
        public string ExtractAccessCode()
        {
            string tokenString = Authorization();
            string accessCode = "";
            if (provider == "yandex.ru")
            {
                //из строки с токеном нужно извлечь только код доступа, создаём объект (класс) tok с нужным полем access_token
                var tok = JsonConvert.DeserializeAnonymousType(token, new { access_token = "" });
                accessCode = tok.access_token;
                Console.WriteLine(accessCode);
            }
            return accessCode;
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
                    using (StreamWriter sw = new StreamWriter("yandex_token.json"))
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
    }
}
