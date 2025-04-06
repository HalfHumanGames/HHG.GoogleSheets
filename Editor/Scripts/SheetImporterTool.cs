using HHG.GoogleSheets.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using UnityEditor;
using UnityEngine;

namespace HHG.GoogleSheets.Editor
{
    public class SheetImporterTool : EditorWindow
    {
        [MenuItem("| Half Human Games|/Tools/Import Google Sheets")]
        public static void ImportSheets()
        {
            foreach (System.Type type in GetAllScriptableObjectSheetTypes())
            {
                SheetAttribute attr = type.GetCustomAttribute<SheetAttribute>();

                if (attr == null)
                {
                    continue;
                }

                Debug.Log($"Importing sheet for {type.Name}: {attr.Gid}");

                string csv = DownloadSheetCSV(attr.SpreadsheetId, attr.Gid);

                if (string.IsNullOrEmpty(csv))
                {
                    Debug.LogError($"Failed to download sheet for {type.Name}");
                    continue;
                }

                ImportCSVToScriptableObjects(csv, type);
            }

            AssetDatabase.SaveAssets();
            Debug.Log("✅ All sheets imported!");
        }

        private static string DownloadSheetCSV(string spreadsheetId, string gid)
        {
            try
            {
                string url = $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/export?format=csv";

                if (!string.IsNullOrEmpty(gid))
                {
                    url += $"&gid={System.Uri.EscapeDataString(gid)}";
                }

                using WebClient client = new WebClient();
                return client.DownloadString(url);
            }
            catch (System.Exception ex)
            {
                Debug.LogError($"Error downloading Google Sheet: {ex.Message}");
                return null;
            }
        }

        private static void ImportCSVToScriptableObjects(string csv, System.Type soType)
        {
            string[] lines = csv.Split(new[] { '\n', '\r' }, System.StringSplitOptions.RemoveEmptyEntries);

            if (lines.Length < 2)
            {
                return;
            }

            string[] headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();

            for (int i = 1; i < lines.Length; i++)
            {
                string[] values = lines[i].Split(',').Select(v => v.Trim()).ToArray();
                Dictionary<string, string> row = headers.Zip(values, (k, v) => new KeyValuePair<string, string>(k, v)).ToDictionary(x => x.Key, x => x.Value);

                if (!row.TryGetValue("Name", out string name) || string.IsNullOrEmpty(name))
                {
                    continue;
                }

                ScriptableObject so = FindScriptableObjectByName(soType, name);

                if (so == null)
                {
                    Debug.LogWarning($"No ScriptableObject found for {soType.Name} named '{name}'");
                    continue;
                }

                ApplyFieldsRecursive(so, row, soType, so);
                EditorUtility.SetDirty(so);
            }
        }

        private static void ApplyFieldsRecursive(object instance, Dictionary<string, string> row, System.Type type, object rootContext)
        {
            foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                // Skip any fields of type UnityEngine.Object
                if (field.FieldType.IsAssignableFrom(typeof(Object)))
                {
                    continue;
                }

                SheetFieldAttribute attr = field.GetCustomAttribute<SheetFieldAttribute>();

                if (attr != null)
                {
                    if (row.TryGetValue(attr.ColumnName, out string raw))
                    {
                        object value = ConvertValue(raw, field.FieldType, attr.TransformMethod, rootContext.GetType());
                        field.SetValue(instance, value);
                    }
                }
                else if (!field.FieldType.IsPrimitive && !field.FieldType.IsEnum && field.FieldType != typeof(string))
                {
                    object subObj = field.GetValue(instance);

                    if (subObj != null)
                    {
                        ApplyFieldsRecursive(subObj, row, field.FieldType, rootContext);
                    }
                }
            }
        }

        private static object ConvertValue(string raw, System.Type targetType, string transformMethod, System.Type contextType)
        {
            if (!string.IsNullOrEmpty(transformMethod))
            {
                MethodInfo method = contextType.GetMethod(transformMethod, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

                if (method != null)
                {
                    return method.Invoke(null, new object[] { raw });
                }
                else
                {
                    Debug.LogWarning($"Transformer method '{transformMethod}' not found on {contextType.Name}.");
                }
            }

            try
            {
                return System.Convert.ChangeType(raw, targetType);
            }
            catch
            {
                Debug.LogWarning($"Failed to convert '{raw}' to {targetType.Name}");
                return targetType.IsValueType ? System.Activator.CreateInstance(targetType) : null;
            }
        }

        private static IEnumerable<System.Type> GetAllScriptableObjectSheetTypes()
        {
            return System.AppDomain.CurrentDomain.GetAssemblies().SelectMany((Assembly a) => a.GetTypes()).Where((System.Type t) => typeof(ScriptableObject).IsAssignableFrom(t) && t.GetCustomAttribute<SheetAttribute>() != null);
        }

        private static ScriptableObject FindScriptableObjectByName(System.Type type, string name)
        {
            string[] guids = AssetDatabase.FindAssets($"t:{type.Name}");

            foreach (string guid in guids)
            {
                string path = AssetDatabase.GUIDToAssetPath(guid);
                ScriptableObject so = AssetDatabase.LoadAssetAtPath(path, type) as ScriptableObject;

                if (so.name == name)
                {
                    return so;
                }
            }

            return null;
        }
    }
}
