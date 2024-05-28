using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnCreateFile_Click(object sender, RoutedEventArgs e)
        {
            //в этом словаре будут храниться тэги и значения. Значения это то, на что будет заменен тег. Тэг есть в договоре, значение береться из БД (т.е. данные сотрудника)
            var items = new Dictionary<string, string>()
            {
                {"<currentYear>", ""},
                {"<currentDay>", ""},
                {"<currentMounth>", ""},
                {"<currentLastTwoCharYear>", "" },
                {"<SotrFio>", "" },
                {"<Role>", "" },
                {"<DayStart>", "" },
                {"<MounthStart>", "" },
                {"<YearStart>", "" },
                {"<salary>", "" },
                {"<Employee>", "" },
                {"<currentDate>", "" }
            };

            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document wordDoc;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();

                object missing = Type.Missing;
                object fileName = ""; //Путь к шаблону документа

                wordDoc = wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing); //открываем шаблон документа

                foreach (var item in items) // Перебор всех тегов и значений словаря, с последующей
                                            // заменой каждого тега на соответствующее для него значение
                {
                    Find find = wordApp.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    object wrap = WdFindWrap.wdFindContinue;
                    object replace = WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false, MatchWholeWord: false, MatchWildcards: false,
                        MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                        Wrap: wrap, Format: false, ReplaceWith: missing, Replace: replace);
                }

                
                // путь по которому будет сохранен файла и имя файла - сохранять на рабочий стол текущего пользователя (или выбор места сохранения через диалоговое окно)
                //
                object newFile = "сюда путь вставить";
                wordDoc.SaveAs2(newFile); //сохранить заполненный данными шаблон как новый документ
                wordApp.ActiveDocument.Close(); //закрытие активного документа
                wordApp?.Quit(); //отключение от приложения для работы с документами типа Word
            }
            catch (Exception ex)
            {
                wordApp.ActiveDocument.Close(); //закрытие активного документа
                wordApp?.Quit();
                Console.WriteLine(ex.Message);
            }

        }
    }
}
