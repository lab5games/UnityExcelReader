using System.IO;
using System.Text;
using UnityEditor;
using UnityEngine;
using UnityEngine.UIElements;

namespace Lab5Games.UnityExcelReader.Editor
{
    public class ExcelToolWindow : EditorWindow
    {
        string _filePath;

        TextField _textFilePath;
        Button _btnSelect;
        Button _btnExportXml;
        Button _btnExportJson;
        Button _btnExportCsv;

        [MenuItem("Lab5Games/Excel Tool Window")]
        public static void ShowExample()
        {
            ExcelToolWindow wnd = GetWindow<ExcelToolWindow>();
            wnd.titleContent = new GUIContent("ExcelToolWindow");
        }

        public void CreateGUI()
        {
            // Each editor window contains a root VisualElement object
            VisualElement root = rootVisualElement;


            // Import UXML
            var visualTree = AssetDatabase.LoadAssetAtPath<VisualTreeAsset>("Assets/UnityExcelReader/Editor/ExcelToolWindow.uxml");

            if(visualTree == null)
                visualTree = AssetDatabase.LoadAssetAtPath<VisualTreeAsset>("Packages/com.lab5games.unityexcelreader/Editor/ExcelToolWindow.uxml");

            //VisualElement labelFromUXML = visualTree.Instantiate();
            //root.Add(labelFromUXML);
            visualTree.CloneTree(root);


            _textFilePath = root.Q<TextField>("textFilePath");

            _btnSelect = root.Q<Button>("btnSelect");
            _btnSelect.clicked += OnClickSelectBtn;

            _btnExportXml = root.Q<Button>("btnExportXml");
            _btnExportXml.clicked += OnClickExportXmlBtn;

            _btnExportJson = root.Q<Button>("btnExportJson");
            _btnExportJson.clicked += OnClickExportJsonBtn;

            _btnExportCsv = root.Q<Button>("btnExportCsv");
            _btnExportCsv.clicked += OnClickExportCsvBtn;
        }

        private void OnClickSelectBtn()
        {
            string path = EditorUtility.OpenFilePanel("Select File", Application.dataPath, "xlsx");

            _filePath = path;
            _textFilePath.value = path;
        }

        private void OnClickExportXmlBtn()
        {
            if (string.IsNullOrEmpty(_filePath))
                return;
        }

        private void OnClickExportJsonBtn()
        {
            if (string.IsNullOrEmpty(_filePath))
                return;
        }

        private void OnClickExportCsvBtn()
        {
            if (string.IsNullOrEmpty(_filePath))
                return;

            StringBuilder csv = new StringBuilder();
            var rows = DataReader.ReadExcelFile(_filePath);

            for(int i=0; i<rows.Count; i++)
            {
                var row = rows[i];

                csv.AppendJoin(',', row.ItemArray);
                csv.AppendLine();
            }

            string dir_name = Path.GetDirectoryName(_filePath);
            string file_namem = Path.GetFileNameWithoutExtension(_filePath) + ".csv";


            File.WriteAllText(Path.Combine(dir_name, file_namem), csv.ToString());

            AssetDatabase.Refresh();

            Debug.Log("csv file created");
        }
    }
}