using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;

namespace RevitMepAddinHost
{
    internal static class Loader
    {
        internal static string LogicDllPath
        {
            get
            {
                string hostDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // 正式安裝：與 RevitAddinHost.dll 放同一層
                string appLocalPath = Path.Combine(hostDir, "RevitMepLogic.dll");
                if (File.Exists(appLocalPath))
                    return appLocalPath;

                // 開發模式 fallback：回專案 dist
                string devDistPath = Path.GetFullPath(
                    Path.Combine(hostDir, @"..\..\..\dist\RevitMepLogic.dll")
                );
                if (File.Exists(devDistPath))
                    return devDistPath;

                // 找不到時，回傳正式路徑，讓錯誤訊息更直觀
                return appLocalPath;
            }
        }

        internal static object Call(string methodPath, UIApplication uiapp = null)
        {
            return Call(methodPath, uiapp == null ? Array.Empty<object>() : new object[] { uiapp });
        }

        internal static object Call(string methodPath, params object[] args)
        {
            if (!File.Exists(LogicDllPath))
                throw new FileNotFoundException("Logic dll not found", LogicDllPath);

            byte[] dllBytes = File.ReadAllBytes(LogicDllPath);

            string pdbPath = Path.ChangeExtension(LogicDllPath, ".pdb");

            Assembly asm = File.Exists(pdbPath)
                ? Assembly.Load(dllBytes, File.ReadAllBytes(pdbPath))
                : Assembly.Load(dllBytes);

            int lastDot = methodPath.LastIndexOf('.');
            if (lastDot < 0)
                throw new ArgumentException("methodPath must include class and method");

            string typeName = methodPath.Substring(0, lastDot);
            string methodName = methodPath.Substring(lastDot + 1);

            Type t = asm.GetType(typeName, throwOnError: true);

            var candidates = t.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance)
                              .Where(x => x.Name == methodName)
                              .ToList();

            if (candidates.Count == 0)
                throw new MissingMethodException(typeName, methodName);

            MethodInfo m = null;
            if (args != null && args.Length > 0)
            {
                var paramTypes = args.Select(a => a?.GetType() ?? typeof(object)).ToArray();
                m = candidates.FirstOrDefault(x =>
                {
                    var ps = x.GetParameters();
                    if (ps.Length != paramTypes.Length) return false;
                    for (int i = 0; i < ps.Length; i++)
                        if (!ps[i].ParameterType.IsAssignableFrom(paramTypes[i])) return false;
                    return true;
                });
            }

            if (m == null)
                m = candidates.FirstOrDefault(x => x.GetParameters().Length == 0)
                    ?? candidates.FirstOrDefault(x =>
                        x.GetParameters().Length == 1 &&
                        x.GetParameters()[0].ParameterType == typeof(UIApplication));

            if (m == null)
                throw new MissingMethodException(typeName, methodName);

            object instance = null;
            object[] invokeArgs = null;

            if (!m.IsStatic)
                instance = Activator.CreateInstance(t);

            if (m.GetParameters().Length > 0)
                invokeArgs = args ?? new object[] { null };

            object result = m.Invoke(instance, invokeArgs);

            string stamp = "";
            try
            {
                var fi = new FileInfo(LogicDllPath);
                stamp = $"{fi.LastWriteTime:HH:mm:ss}";
            }
            catch { }

            return $"[LogicDll:{stamp}] {result}";
        }
    }
}