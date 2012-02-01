//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2010  Quizo, Paul Accisano
//
//    QTTabBar is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    QTTabBar is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with QTTabBar.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceModel;
using System.Threading;
using System.Windows.Threading;
using QTTabBarLib.Interop;

namespace QTTabBarLib {
    internal static class InstanceManager {
        private static Dictionary<Thread, QTTabBarClass> dictTabInstances = new Dictionary<Thread, QTTabBarClass>();
        private static Dictionary<Thread, QTButtonBar> dictBBarInstances = new Dictionary<Thread, QTButtonBar>();
        private static StackDictionary<IntPtr, QTTabBarClass> sdTabHandles = new StackDictionary<IntPtr, QTTabBarClass>();
        private static ReaderWriterLock rwLockBtnBar = new ReaderWriterLock();
        private static ReaderWriterLock rwLockTabBar = new ReaderWriterLock();
        private static Dispatcher commDispatch;
        private static IServiceContract commChannel;
        private static ICallbackContract commCallback;
        private static bool isServer;

        // Server stuff
        private static ServiceHost serviceHost;
        private static List<ICallbackContract> callbacks = new List<ICallbackContract>();
        private static StackDictionary<IntPtr, ICallbackContract> sdInstances = new StackDictionary<IntPtr, ICallbackContract>();

        // Client stuff;
        private static DuplexChannelFactory<IServiceContract> pipeFactory;

        #region Comm Classes / Interfaces

        [ServiceContract(SessionMode = SessionMode.Required, CallbackContract = typeof(ICallbackContract))]
        private interface IServiceContract {
            [OperationContract]
            void Subscribe();

            [OperationContract]
            void PushInstance(IntPtr hwnd);

            [OperationContract]
            void DeleteInstance(IntPtr hwnd);

            [OperationContract]
            bool IsMainProcess();

            [OperationContract]
            int GetTotalInstanceCount();

            [OperationContract]
            bool ExecuteOnMainProcess(byte[] encodedAction);

            [OperationContract]
            void Broadcast(byte[] encodedAction);
        }

        [ServiceBehavior(
                ConcurrencyMode = ConcurrencyMode.Reentrant,
                InstanceContextMode = InstanceContextMode.PerSession)]
        private class CommServer : IServiceContract {

            private static void CheckConnections() {
                callbacks.RemoveAll(callback => {
                    ICommunicationObject ico = callback as ICommunicationObject;
                    return ico != null && ico.State != CommunicationState.Opened;
                });
                sdInstances.RemoveAllValues(c => !callbacks.Contains(c));
            }

            private static ICallbackContract GetCallback() {
                OperationContext context = OperationContext.Current;
                return context == null
                        ? commCallback
                        : context.GetCallbackChannel<ICallbackContract>();
            }

            public int GetTotalInstanceCount() {
                lock(callbacks) {
                    CheckConnections();
                    return sdInstances.Count;
                }
            }

            public bool ExecuteOnMainProcess(byte[] encodedAction) {
                ICallbackContract callback;
                lock(callbacks) {
                    CheckConnections();
                    if(IsMainProcess()) {
                        return true;
                    }
                    else if(sdInstances.Count == 0) {
                        return false;
                    }
                    callback = sdInstances.Peek();
                }
                callback.Execute(encodedAction);
                return false;
            }

            public void Broadcast(byte[] encodedAction) {
                ICallbackContract sender = GetCallback();
                Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() => {
                    lock(callbacks) {
                        CheckConnections();
                        foreach(ICallbackContract callback in callbacks) {
                            if(callback != sender) {
                                callback.Execute(encodedAction);
                            }
                        }
                    }
                }), DispatcherPriority.Normal);
            }

            public void DeleteInstance(IntPtr hwnd) {
                lock(callbacks) {
                    CheckConnections();
                    sdInstances.Remove(hwnd);
                    // todo: confirm liveness of new stack head
                }
            }

            public bool IsMainProcess() {
                lock(callbacks) {
                    CheckConnections();
                    return sdInstances.Count > 0 && GetCallback() == sdInstances.Peek();
                }
            }

            public void Subscribe() {
                lock(callbacks) {
                    ICallbackContract callback = GetCallback();
                    if(!callbacks.Contains(callback)) {
                        callbacks.Add(callback);
                    }
                }
            }

            public void PushInstance(IntPtr hwnd) {
                lock(callbacks) {
                    CheckConnections();
                    if(!callbacks.Contains(GetCallback())) return; // hmmm....
                    sdInstances.Push(hwnd, GetCallback());
                }
            }
        }

        private interface ICallbackContract {
            [OperationContract]
            void Execute(byte[] encodedAction);
        }

        [CallbackBehavior(ConcurrencyMode = ConcurrencyMode.Reentrant)]
        private class CommClient : ICallbackContract {
            public void Execute(byte[] encodedAction) {
                try {
                    ByteToDel(encodedAction).DynamicInvoke();
                }
                catch(Exception ex) {
                    QTUtility2.MakeErrorLog(ex);
                }
            }
        }

        #endregion

        #region Utility Methods

        private static byte[] DelToByte(Delegate del) {
            return QTUtility.ObjectToByteArray(new SerializeDelegate(del));
        }

        private static Delegate ByteToDel(byte[] buf) {
            return ((SerializeDelegate)QTUtility.ByteArrayToObject(buf)).Delegate;
        }

        private static void CommBeginInvoke(Action action) {
            commDispatch.BeginInvoke(action, DispatcherPriority.Normal);
        }

        private static void CommInvoke(Action action) {
            commDispatch.Invoke(action);
        }

        #endregion

        public static void Initialize() {
            Thread thread = new Thread(Dispatcher.Run) {IsBackground = true};
            thread.Start();
            while(true) {
                commDispatch = Dispatcher.FromThread(thread);
                if(commDispatch != null) break;
                Thread.Sleep(50);
            }
            commCallback = new CommClient();

            uint desktopPID;
            PInvoke.GetWindowThreadProcessId(WindowUtils.GetShellTrayWnd(), out desktopPID);
            isServer = desktopPID == PInvoke.GetCurrentProcessId();
            const string PipeName = "QTTabBarPipe";
            string address = "net.pipe://localhost/" + PipeName + desktopPID;
            CommInvoke(() => {
                if(isServer) {
                    commChannel = new CommServer();
                    serviceHost = new ServiceHost(
                            typeof(CommServer),
                            new Uri[] { new Uri(address) });
                    serviceHost.AddServiceEndpoint(
                            typeof(IServiceContract),                                 // TODO: this is only 24 hours...
                            new NetNamedPipeBinding(NetNamedPipeSecurityMode.None) { ReceiveTimeout = TimeSpan.MaxValue },
                            new Uri(address));
                    serviceHost.Open();
                    commChannel.Subscribe();
                }
                else {
                    pipeFactory = new DuplexChannelFactory<IServiceContract>(
                            commCallback,
                            new NetNamedPipeBinding(NetNamedPipeSecurityMode.None),
                            new EndpointAddress(address));
                    commChannel = pipeFactory.CreateChannel();
                    try {
                        commChannel.Subscribe();
                    }
                    catch(EndpointNotFoundException) {
                        // todo: ???
                    }
                }
            });
        }

        public static void StaticBroadcast(Action action) {
            commChannel.Broadcast(DelToByte(action));
        }

        public static void TabBarBroadcast(Action<QTTabBarClass> action) {
            LocalTabBroadcast(action, Thread.CurrentThread);
            StaticBroadcast(() => LocalTabBroadcast(action));
        }

        private static void LocalTabBroadcast(Action<QTTabBarClass> action, Thread skip = null) {
            using(new Keychain(rwLockTabBar, false)) {
                foreach(var pair in dictTabInstances) {
                    if(pair.Key != skip) {
                        pair.Value.Invoke(action, pair.Value);   
                    }
                }
            }
        }

        public static void ButtonBarBroadcast(Action<QTButtonBar> action) {
            LocalBBarBroadcast(action, Thread.CurrentThread);
            StaticBroadcast(() => LocalBBarBroadcast(action));
        }

        private static void LocalBBarBroadcast(Action<QTButtonBar> action, Thread skip = null) {
            using(new Keychain(rwLockBtnBar, false)) {
                foreach(var pair in dictBBarInstances) {
                    if(pair.Key != skip) {
                        pair.Value.Invoke(action, pair.Value);
                    }
                }
            }
        }

        public static bool EnsureMainProcess(Action action) {
            if(commChannel.IsMainProcess()) return true;
            ExecuteOnMainProcess(action);
            return false;
        }

        private static void ExecuteOnMainProcess(Action action) {
            if(commChannel.ExecuteOnMainProcess(DelToByte(action))) {
                action();
            }
        }

        public static void InvokeMain(Action<QTTabBarClass> action) {
            ExecuteOnMainProcess(() => LocalInvokeMain(action));
        }

        private static void LocalInvokeMain(Action<QTTabBarClass> action) {
            QTTabBarClass instance;
            using(new Keychain(rwLockTabBar, false)) {
                instance = sdTabHandles.Count == 0 ? null : sdTabHandles.Peek();
            }
            if(instance != null) instance.Invoke(action, instance);
        }

        public static void RegisterButtonBar(QTButtonBar bbar) {
            using(new Keychain(rwLockBtnBar, true)) {
                dictBBarInstances[Thread.CurrentThread] = bbar;
            }
        }

        public static QTTabBarClass GetThreadTabBar() {
            using(new Keychain(rwLockTabBar, false)) {
                QTTabBarClass tab;
                return dictTabInstances.TryGetValue(Thread.CurrentThread, out tab) ? tab : null;
            }
        }

        public static void RegisterTabBar(QTTabBarClass tabbar) {
            IntPtr handle = tabbar.Handle;
            using(new Keychain(rwLockTabBar, true)) {
                dictTabInstances[Thread.CurrentThread] = tabbar;
                sdTabHandles.Push(handle, tabbar);
            }
            //CommBeginInvoke(() => commChannel.PushInstance(handle));
            commChannel.PushInstance(handle);
        }

        public static void UnregisterButtonBar() {
            using(new Keychain(rwLockBtnBar, true)) {
                dictBBarInstances.Remove(Thread.CurrentThread);
            }
        }

        public static bool UnregisterTabBar() {
            using(new Keychain(rwLockTabBar, true)) {
                QTTabBarClass tabbar;
                if(dictTabInstances.TryGetValue(Thread.CurrentThread, out tabbar)) {
                    IntPtr handle = tabbar.Handle;
                    dictTabInstances.Remove(Thread.CurrentThread);
                    sdTabHandles.Remove(handle);
                    CommBeginInvoke(() => commChannel.DeleteInstance(handle));
                }
                return false;
            }
        }

        public static bool TryGetButtonBarHandle(IntPtr explorerHandle, out IntPtr ptr) {
            // todo
            QTButtonBar bbar;
            if(dictBBarInstances.TryGetValue(Thread.CurrentThread, out bbar)) {
                ptr = bbar.Handle;
                return true;
            }
            ptr = IntPtr.Zero;
            return false;
        }

        public static int GetTotalInstanceCount() {
            return commChannel.GetTotalInstanceCount();
        }
    }
}
