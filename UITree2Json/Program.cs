using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using Newtonsoft.Json;

namespace UITree2Json
{
    class WindowsInterop
    {
        internal delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn,
            IntPtr lParam);
    }

    class Program
    {
        static object ElementAndDescendents2SerializableObject(
            AutomationElement rootElement,
            int maxDepth,
            List<string> ignoreChildrenOfTheseClassNames)
        {
            Dictionary<string, object> properties = new Dictionary<string, object>
            {
                {"name", rootElement.Current.Name},
                {"className", rootElement.Current.ClassName},
            };

            if(!ignoreChildrenOfTheseClassNames.Any(
                stopDiggingAfterThisClassName => rootElement.Current.ClassName == stopDiggingAfterThisClassName
                ))
            {
                List<object> children = new List<object>();
                if (maxDepth > 0)
                {
                    AutomationElement child = TreeWalker.ControlViewWalker.GetFirstChild(rootElement);
                    while (child != null)
                    {
                        children.Add(ElementAndDescendents2SerializableObject(child, maxDepth - 1, ignoreChildrenOfTheseClassNames));
                        child = TreeWalker.ControlViewWalker.GetNextSibling(child);
                    }
                }

                if (children.Count > 0)
                {
                    properties.Add("children", children);
                }
            }

            return properties;
        }

        static void Main(string[] args)
        {
            if(args.Any(arg => arg.Contains("?")))
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("UITree2Json.exe [-maxdepth <n>] [-ignorebeneath <classname>] [processname0] [processname1]");
                Console.WriteLine();
                Console.WriteLine("If you list no processes, I'll print the entire UI tree. Hope that's what you wanted!");
            }
            else if (args.Length > 0)
            {
                int maxDepth = int.MaxValue;
                List<string> ignoreChildrenOfTheseClassNames = new List<string>();

                List<string> prunedArgs = new List<string>(args);

                int maxDepthFlagIndex = prunedArgs.FindIndex(arg => arg.Equals("-maxdepth", StringComparison.OrdinalIgnoreCase));
                if(maxDepthFlagIndex != -1)
                {
                    maxDepth = int.Parse(args[maxDepthFlagIndex + 1]);
                    prunedArgs.RemoveRange(maxDepthFlagIndex, 2);
                }

                while (true)
                {
                    int ignoreChildrenFlagIndex = prunedArgs.FindIndex(arg => arg.Equals("-ignorebeneath", StringComparison.OrdinalIgnoreCase));
                    if (ignoreChildrenFlagIndex == -1)
                    {
                        break;
                    }

                    ignoreChildrenOfTheseClassNames.Add(prunedArgs[ignoreChildrenFlagIndex + 1]);
                    prunedArgs.RemoveRange(ignoreChildrenFlagIndex, 2);
                }

                prunedArgs.Select(
                    providedProcessName => Regex.Replace(providedProcessName, "\\.exe$", "", RegexOptions.IgnoreCase)
                    ).SelectMany(
                    processName => Process.GetProcessesByName(processName)
                    ).SelectMany(
                    process => process.Threads.Cast<ProcessThread>()
                    ).ToList().SelectMany(
                    thread => { 
                        List<IntPtr> hwnds = new List<IntPtr>();
                        WindowsInterop.EnumThreadWindows(thread.Id,
                            (hWnd, lParam) => { hwnds.Add(hWnd); return true; },
                            IntPtr.Zero);
                        return hwnds;
                    }).Select(
                    hwnd => AutomationElement.FromHandle(hwnd)
                    ).ToList().ForEach(
                    element => Console.WriteLine(JsonConvert.SerializeObject(
                        ElementAndDescendents2SerializableObject(element, maxDepth, ignoreChildrenOfTheseClassNames)))
                        );
            }
            else
            {
                Console.WriteLine(JsonConvert.SerializeObject(
                    ElementAndDescendents2SerializableObject(AutomationElement.RootElement, maxDepth: int.MaxValue, ignoreChildrenOfTheseClassNames: new List<string>())
                ));
            }
        }
    }
}
