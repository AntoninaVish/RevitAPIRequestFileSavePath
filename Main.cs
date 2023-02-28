using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIRequestFileSavePath
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //создаем переменную, которая будет собирать данные по помещению
            string roomInfo = string.Empty;//переменная в которую эти данные будут записываться

            var rooms = new FilteredElementCollector(doc) //собираем сами помещения с помощью FilteredElementCollector
            .OfCategory(BuiltInCategory.OST_Rooms) //собираем по категории
            .Cast<Room>() //элементы, которые получаем преобразовываем в помещения
            .ToList();//преобразовываем в список

            //заполняем данный параметр roominfo, проходимся в цикле по каждому помещению
            foreach (Room room in rooms)
            {
                //извлекаем имя помещения, для этого создаем отдельную переменную, с помощью get_parameter заходим
                //в строенный параметри находим встроенный параметр ROOM_NAME, AsString - имя помещения выводится в строку
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                //заполняем переменную roomInfo новыми данными, \t - это знак табуляции, чтобы выводилось в разном формате txt, xl
                roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}";//имя, номер, площадь и все выводиться
                                                                                           //с новой строки
            }

            //вызвать путь сохранения файла понадобится переменная типа saveDialog
            var saveDialog = new SaveFileDialog
            {
                
                OverwritePrompt = true, //заполняем свойства, которые будут необходимы, OverwritePrompt
                                        //выдает запрос на перезапись файла
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),//какая папка будет выдаваться
                                                                                               //по умолчанию, напр., папка рабочего стола
                Filter = "All files (*.*)|*.*", //указываем фильтр, если хотим чтобы отображались все файлы то прописываем; 
                                             //*.* -это значит что при сохранении файла в папке будут отображаться все файлы с любым расширением
                FileName = "roomInfo.csv",//название файла по умолчанию, при этом название файла можно будет менять при сохранении
                DefaultExt = ".csv" //расширение по умолчанию

            };
            //создаем переменную, которая сохранит выбраный пользователем путь, строка будет пустой
            string selectedFilePath = string.Empty;

            //если сохранение файла прошло успешно, тогда забераем введенный путь и сохраняем его в переменную selectedFilePath
            if(saveDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveDialog.FileName;
            }
            //если путь не был указан, то выходим из данного метода и завершаем выполнение программы
            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            //если все впорядке путь у нас есть тогда вызываем сттический метод WriteAllText
            File.WriteAllText(selectedFilePath, roomInfo);
            
            
            return Result.Succeeded;
        }
    }
}
