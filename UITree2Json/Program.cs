using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
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
            List<string> ignoreChildrenOfTheseClassNames,
            bool addStructureChangedEventHandler = false)
        {
            try
            {
                if (addStructureChangedEventHandler)
                {
                    Automation.AddStructureChangedEventHandler(rootElement, TreeScope.Descendants,
                        (sender, e) =>
                        {
                            AutomationElement senderElement = (AutomationElement)sender;
                            if(senderElement != null)
                            {
                                // Check if this event is coming from below a class that we're ignoring
                                bool ignoreThisOne = false;
                                try
                                {
                                    AutomationElement ancestor = TreeWalker.ControlViewWalker.GetParent(senderElement);
                                    while (ancestor != null)
                                    {
                                        string cachedAncestorClassName = ancestor.Current.ClassName;
                                        if (ignoreChildrenOfTheseClassNames.Any(
                                            stopDiggingAfterThisClassName => cachedAncestorClassName == stopDiggingAfterThisClassName
                                            ))
                                        {
                                            ignoreThisOne = true;
                                            break;
                                        }

                                        ancestor = TreeWalker.ControlViewWalker.GetParent(ancestor);
                                    }
                                }
                                catch (ElementNotAvailableException)
                                {
                                }

                                if (!ignoreThisOne)
                                {
                                    SerializeToConsoleIfNotNull(
                                        ElementAndDescendents2SerializableObject(senderElement, maxDepth, ignoreChildrenOfTheseClassNames)
                                        );
                                }
                            }
                        });
                }

                string cachedClassName = rootElement.Current.ClassName;
                Dictionary<string, object> properties = new Dictionary<string, object>
                {
                    {"name", rootElement.Current.Name},
                    {"className", cachedClassName},
                    {"isOffscreen", rootElement.Current.IsOffscreen},
                };

                if (!ignoreChildrenOfTheseClassNames.Any(
                    stopDiggingAfterThisClassName => cachedClassName == stopDiggingAfterThisClassName
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
            catch (ElementNotAvailableException)
            {
                // It's gone
                return null;
            }
        }

        static void SerializeToConsoleIfNotNull(object obj)
        {
            if (obj != null)
            {
                Console.WriteLine(JsonConvert.SerializeObject(obj));
            }
        }

        static void Main(string[] args)
        {
            if (args.Any(arg => arg.Contains("?")))
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("UITree2Json.exe [-maxdepth <n>] [-ignorebeneath <classname>] [-subscribe2updates] [processname0] [processname1]");
                Console.WriteLine();
                Console.WriteLine("If you list no processes, I'll print the entire UI tree. Hope that's what you wanted!");
                Console.WriteLine("subscribe2updates is only available if you specify process names");
            }
            else if (args.Length > 0)
            {
                List<string> prunedArgs = new List<string>(args);

                ManualResetEvent stayRunningUntilThisHandleIsSignalled = null;
                {
                    int subscribe2UpdatesSwitchIndex = prunedArgs.FindIndex(arg => arg.Equals("-subscribe2updates", StringComparison.OrdinalIgnoreCase));
                    if (subscribe2UpdatesSwitchIndex != -1)
                    {
                        stayRunningUntilThisHandleIsSignalled = new ManualResetEvent(false);
                        Console.CancelKeyPress += (sender, e) =>
                        {
                            stayRunningUntilThisHandleIsSignalled.Set();
                        };

                        prunedArgs.RemoveAt(subscribe2UpdatesSwitchIndex);
                    }
                }

                int maxDepth = int.MaxValue;
                {
                    int maxDepthOptionIndex = prunedArgs.FindIndex(arg => arg.Equals("-maxdepth", StringComparison.OrdinalIgnoreCase));
                    if (maxDepthOptionIndex != -1)
                    {
                        maxDepth = int.Parse(args[maxDepthOptionIndex + 1]);
                        prunedArgs.RemoveRange(maxDepthOptionIndex, 2);
                    }
                }

                List<string> ignoreChildrenOfTheseClassNames = new List<string>();
                while (true)
                {
                    int ignoreChildrenOptionIndex = prunedArgs.FindIndex(arg => arg.Equals("-ignorebeneath", StringComparison.OrdinalIgnoreCase));
                    if (ignoreChildrenOptionIndex == -1)
                    {
                        break;
                    }

                    ignoreChildrenOfTheseClassNames.Add(prunedArgs[ignoreChildrenOptionIndex + 1]);
                    prunedArgs.RemoveRange(ignoreChildrenOptionIndex, 2);
                }

                var unrecognizedSwitches = prunedArgs.Where(
                    arg => arg.StartsWith("-")
                    );
                if (unrecognizedSwitches.Count() > 0)
                {
                    unrecognizedSwitches.ToList().ForEach(
                        arg => Console.WriteLine("Unrecognized command line option: {0}", arg)
                        );
                }
                else
                {
                    var processesWeCareAbout = prunedArgs.Select(
                        providedProcessName => Regex.Replace(providedProcessName, "\\.exe$", "", RegexOptions.IgnoreCase)
                        );

                    if (stayRunningUntilThisHandleIsSignalled != null)
                    {
                        Automation.AddAutomationEventHandler(
                            WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement,
                            TreeScope.Descendants,
                            (sender, e) =>
                            {
                                try
                                {
                                    AutomationElement sourceElement = (AutomationElement)sender;
                                    Process p = Process.GetProcessById(sourceElement.Current.ProcessId);
                                    if (sourceElement != null &&
                                        processesWeCareAbout.Any(
                                        processName => p.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase)
                                        ))
                                    {
                                        SerializeToConsoleIfNotNull(
                                            ElementAndDescendents2SerializableObject(sourceElement, maxDepth,
                                                ignoreChildrenOfTheseClassNames, addStructureChangedEventHandler: true)
                                            );
                                    }
                                }
                                catch (ElementNotAvailableException)
                                {
                                    // It might've already been destroyed
                                }
                            }
                            );
                    }

                    processesWeCareAbout.SelectMany(
                        processName => Process.GetProcessesByName(processName)
                        ).SelectMany(
                        process => process.Threads.Cast<ProcessThread>()
                        ).ToList().SelectMany(
                        thread =>
                        {
                            List<IntPtr> hwnds = new List<IntPtr>();
                            WindowsInterop.EnumThreadWindows(thread.Id,
                                (hWnd, lParam) => { hwnds.Add(hWnd); return true; },
                                IntPtr.Zero);
                            return hwnds;
                        }).Select(
                        hwnd => AutomationElement.FromHandle(hwnd)
                        ).ToList().ForEach(
                        element => SerializeToConsoleIfNotNull(
                            ElementAndDescendents2SerializableObject(element, maxDepth, ignoreChildrenOfTheseClassNames,
                                addStructureChangedEventHandler: stayRunningUntilThisHandleIsSignalled != null))
                            );

                    if (stayRunningUntilThisHandleIsSignalled != null)
                    {
                        stayRunningUntilThisHandleIsSignalled.WaitOne();
                    }
                }
            }
            else
            {
                SerializeToConsoleIfNotNull(
                    ElementAndDescendents2SerializableObject(AutomationElement.RootElement, maxDepth: int.MaxValue, ignoreChildrenOfTheseClassNames: new List<string>())
                );
            }
        }
    }
}
