using HHG.Common.Runtime;
using HHG.GoogleSheets.Runtime;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;
using UnityEditor;
using UnityEngine;

namespace HHG.GoogleSheets.Editor
{
    public class SheetImporterTool : EditorWindow
    {
        private static HashSet<int> visited = new HashSet<int>();
        private static HashSet<System.Type> loaded = new HashSet<System.Type>();
        private static Dictionary<string, ScriptableObject> cache = new Dictionary<string, ScriptableObject>();

        private struct Sheet
        {
            public string SpreadsheetId;
            public string Gid;

            public Sheet(string spreadsheetId, string gid)
            {
                SpreadsheetId = spreadsheetId;
                Gid = gid;
            }

            public override bool Equals(object other)
            {
                return other is Sheet import && SpreadsheetId == import.SpreadsheetId && Gid == import.Gid;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + (SpreadsheetId != null ? SpreadsheetId.GetHashCode() : 0);
                    hash = hash * 23 + (Gid != null ? Gid.GetHashCode() : 0);
                    return hash;
                }
            }
        }

        [MenuItem("| Half Human Games |/Tools/Import Selected Google Sheet(s)")]
        public static void ImportSelectedSheets()
        {
            ImportSheetHelper(GetSelectedScriptableObjectSheetTypes());
        }

        [MenuItem("| Half Human Games |/Tools/Import All Google Sheets")]
        public static void ImportAllSheets()
        {
            ImportSheetHelper(GetAllScriptableObjectSheetTypes());
        }

        private static void ImportSheetHelper(IEnumerable<System.Type> types)
        {
            loaded.Clear();
            cache.Clear();

            HashSet<Sheet> imported = new HashSet<Sheet>();

            foreach (System.Type type in types)
            {
                SheetAttribute attr = type.GetCustomAttribute<SheetAttribute>();

                if (attr == null)
                {
                    continue;
                }

                Debug.Log($"Importing sheet for {type.Name}");

                string csv = DownloadSheetCSV(attr.SpreadsheetId, attr.Gid);

                if (string.IsNullOrEmpty(csv))
                {
                    Debug.LogError($"Failed to download sheet for {type.Name}");
                    continue;
                }

                ImportCSVToScriptableObjects(csv, type, attr.Casing);

                imported.Add(new Sheet(attr.SpreadsheetId, attr.Gid));
            }

            // Save Assets vefore invoke callbacks
            AssetDatabase.SaveAssets();

            foreach (Sheet sheet in imported)
            {
                InvokeSheetImportedCallbacks(sheet.SpreadsheetId, sheet.Gid);
            }

            Debug.Log("All Google Sheets imported!");
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

        private static void ImportCSVToScriptableObjects(string csv, System.Type type, Case casing)
        {
            List<List<string>> rows = CsvUtil.Parse(csv);

            if (rows.Count < 2)
            {
                return;
            }

            List<string> headers = rows[0];
            Dictionary<string, string> row = new Dictionary<string, string>();
            Dictionary<string, int> arrayTracker = new Dictionary<string, int>();

            for (int i = 1; i < rows.Count; i++)
            {
                row.Clear();
                row.AddRange(headers.Zip(rows[i], (k, v) => new KeyValuePair<string, string>(k, v)));
                arrayTracker.Clear();

                if (!row.TryGetValue("Name", out string name) || string.IsNullOrEmpty(name))
                {
                    continue;
                }

                ScriptableObject scriptableObject = FindScriptableObjectByName(type, name);

                if (scriptableObject == null)
                {
                    Debug.LogWarning($"No {type.Name} named '{name}' found.");
                    continue;
                }

                visited.Clear();
                ApplyFieldsRecursive(scriptableObject, row, type, scriptableObject, casing, arrayTracker);
                EditorUtility.SetDirty(scriptableObject);
            }
        }

        private static void ApplyFieldsRecursive(object instance, Dictionary<string, string> row, System.Type type, object rootContext, Case casing, Dictionary<string, int> arrayTracker)
        {
            // Track visited fields to prevent stack overflow
            // caused from an indefinite recursion loop
            int hash = RuntimeHelpers.GetHashCode(instance);
            if (!visited.Add(hash)) return;

            foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                SheetFieldAttribute attr = field.GetCustomAttribute<SheetFieldAttribute>();

                if (attr != null)
                {
                    string columnName = string.IsNullOrEmpty(attr.ColumnName) ? casing.ToCase(field.Name) : attr.ColumnName;
                    int index = arrayTracker.GetValueOrDefault(columnName);
                    bool increment = false;

                    if (row.TryGetValue(columnName, out string contents) ||
                       (increment = row.TryGetValue($"{columnName} {index}", out contents)))
                    {
                        object value = ConvertValue(columnName, contents, field.FieldType, attr.TransformMethod, rootContext.GetType());
                        SetValue(columnName, contents, field.FieldType, attr.AdapterMethod, rootContext.GetType(), field, instance, value);

                        if (increment)
                        {
                            arrayTracker[columnName] = index + 1;
                        }
                    }
                }
                else if (IsValidFieldType(field.FieldType) && field.GetValue(instance) is object obj)
                {
                    if (obj is IEnumerable enumerable && (field.FieldType.GetElementType() ?? field.FieldType.GetGenericArguments().FirstOrDefault()) != null)
                    {
                        foreach (object item in enumerable)
                        {
                            if (item is null)
                            {
                                continue;
                            }

                            ApplyFieldsRecursive(item, row, item.GetType(), rootContext, casing, arrayTracker);
                        }
                    }
                    else
                    {
                        ApplyFieldsRecursive(obj, row, field.FieldType, rootContext, casing, arrayTracker);
                    }
                }
            }
        }

        private static object ConvertValue(string columnName, string input, System.Type targetType, string transformMethod, System.Type contextType)
        {
            MethodInfo method = !string.IsNullOrEmpty(transformMethod) ?
                contextType.GetMethod(transformMethod, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic) :
                FindMethodByColumnName<SheetTransformAttribute>(contextType, columnName) ?? FindMethodByFieldType<SheetTransformAttribute>(contextType, targetType);

            if (method != null)
            {
                try
                {
                    return method.Invoke(null, new object[] { input });
                }
                catch (System.Exception e)
                {
                    Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                    return targetType.IsValueType ? System.Activator.CreateInstance(targetType) : null;
                }
            }
            else
            {
                try
                {
                    return System.Convert.ChangeType(input, targetType);
                }
                catch (System.Exception e)
                {
                    Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                    return targetType.IsValueType ? System.Activator.CreateInstance(targetType) : null;
                }
            }
        }

        private static void SetValue(string columnName, string input, System.Type targetType, string adapterMethod, System.Type contextType, FieldInfo field, object instance, object value)
        {
            MethodInfo method = !string.IsNullOrEmpty(adapterMethod) ?
                contextType.GetMethod(adapterMethod, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic) :
                FindMethodByColumnName<SheetAdapterAttribute>(contextType, columnName) ?? FindMethodByFieldType<SheetAdapterAttribute>(contextType, targetType);

            if (method != null)
            {
                try
                {
                    method.Invoke(null, new object[] { instance, value });
                }
                catch (System.Exception e)
                {
                    Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                }
            }
            else
            {
                try
                {
                    field.SetValue(instance, value);
                }
                catch (System.Exception e)
                {
                    Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                }
            }
        }

        private static MethodInfo FindMethodByColumnName<T>(System.Type contextType, string columnName) where T : System.Attribute, ISheetMethod
        {
            return FindMethod<T>(contextType, attr => attr.ColumnName == columnName);
        }

        private static MethodInfo FindMethodByFieldType<T>(System.Type contextType, System.Type targetType) where T : System.Attribute, ISheetMethod
        {
            return FindMethod<T>(contextType, attr => attr.FieldType == targetType);
        }

        private static MethodInfo FindMethod<T>(System.Type contextType, System.Func<T, bool> predicate) where T : System.Attribute, ISheetMethod
        {
            // Make sure to flatten to also include base methods
            MethodInfo[] methods = contextType.GetMethods(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.FlattenHierarchy);

            MethodInfo method = methods.FirstOrDefault(m =>
            {
                T attribute = m.GetCustomAttribute<T>();
                return attribute != null && predicate(attribute);
            });

            if (method != null)
            {
                return method;
            }

            // No need to flatten the fields since they use recursion
            // intead since it's more performant in this scenario
            foreach (FieldInfo field in contextType.GetFields(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public))
            {
                if (!IsValidFieldType(field.FieldType))
                {
                    continue;
                }

                method = FindMethod(field.FieldType, predicate);

                if (method != null)
                {
                    return method;
                }
            }

            return null;
        }

        private static bool IsValidFieldType(System.Type fieldType)
        {
            return !fieldType.IsPrimitive && !fieldType.IsEnum && fieldType != typeof(string) && !typeof(Object).IsAssignableFrom(fieldType) &&
                  ((fieldType.GetElementType() ?? fieldType.GetGenericArguments().FirstOrDefault()) is not System.Type subType || IsValidFieldType(subType));
        }

        private static IEnumerable<System.Type> GetSelectedScriptableObjectSheetTypes()
        {
            return Selection.objects.OfType<ScriptableObject>().Select(so => so.GetType()).Distinct().Where(t => t.GetCustomAttribute<SheetAttribute>() != null);
        }

        private static IEnumerable<System.Type> GetAllScriptableObjectSheetTypes()
        {
            return System.AppDomain.CurrentDomain.GetAssemblies().SelectMany((Assembly a) => a.GetTypes()).Where((System.Type t) => typeof(ScriptableObject).IsAssignableFrom(t) && t.GetCustomAttribute<SheetAttribute>() != null);
        }

        private static ScriptableObject FindScriptableObjectByName(System.Type type, string name)
        {
            if (!loaded.Contains(type))
            {
                loaded.Add(type);

                foreach (string guid in AssetDatabase.FindAssets($"t:{type.Name}"))
                {
                    string path = AssetDatabase.GUIDToAssetPath(guid);
                    ScriptableObject scriptableObject = AssetDatabase.LoadAssetAtPath(path, type) as ScriptableObject;
                    cache[scriptableObject.name] = scriptableObject;
                }
            }

            if (!cache.ContainsKey(name)) cache[name] = null;

            return cache[name];
        }

        private static void InvokeSheetImportedCallbacks(string spreadsheetId, string gid)
        {
            foreach (System.Type type in System.AppDomain.CurrentDomain.GetAssemblies().SelectMany(a => a.GetTypes()))
            {
                foreach (MethodInfo method in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    SheetImportedCallbackAttribute attr = method.GetCustomAttribute<SheetImportedCallbackAttribute>();

                    if (attr != null && 
                        (string.IsNullOrEmpty(attr.SpreadsheetId) || attr.SpreadsheetId == spreadsheetId) && 
                        (attr.Gids == null || attr.Gids.Length == 0 || attr.Gids.IndexOf(gid) >= 0))
                    {
                        method.Invoke(null, null);
                    }
                }
            }
        }
    }
}
