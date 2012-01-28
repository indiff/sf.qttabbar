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
using System.Security.Principal;
using System.ServiceModel;
using System.Threading;
using QTTabBarService;

namespace QTTabBarLib {
    internal static class InstanceManager {
        private static Dictionary<Thread, QTTabBarClass> dictTabInstances = new Dictionary<Thread, QTTabBarClass>();
        private static Dictionary<Thread, QTButtonBar> dictBBarInstances = new Dictionary<Thread, QTButtonBar>();
        private static Dictionary<IntPtr, QTTabBarClass> dictTabHandles = new Dictionary<IntPtr, QTTabBarClass>();
        private static ReaderWriterLock rwLockBtnBar = new ReaderWriterLock();
        private static ReaderWriterLock rwLockTabBar = new ReaderWriterLock();
        private static QTTabBarClass mainInstance = null; // TODO

        private static DuplexChannelFactory<IServiceContract> pipeFactory;
        private static IServiceContract commChannel;

        private class Keychain : IDisposable {
            private ReaderWriterLock rwlock;
            private bool write;

            public Keychain(ReaderWriterLock rwlock, bool write) {
                this.rwlock = rwlock;
                this.write = write;
                if(write) {
                    rwlock.AcquireWriterLock(Timeout.Infinite);
                }
                else {
                    rwlock.AcquireReaderLock(Timeout.Infinite);
                }
            }

            public void Dispose() {
                if(rwlock == null) return;
                if(write) {
                    rwlock.ReleaseWriterLock();
                }
                else {
                    rwlock.ReleaseReaderLock();
                }
                rwlock = null;
            }
        }

        [CallbackBehavior(UseSynchronizationContext = false)]
        private class CommCallback : ICallbackContract {
            public bool SetMain(IntPtr hwnd) {
                using(new Keychain(rwLockTabBar, true)) {
                    if(dictTabHandles.ContainsKey(hwnd)) {
                        mainInstance = dictTabHandles[hwnd];
                        return true;
                    }
                    return false;
                }
            }

            public void Execute(byte[] encodedAction) {
                SerializeDelegate sd = (SerializeDelegate)QTUtility.ByteArrayToObject(encodedAction);
                sd.Delegate.DynamicInvoke();
            }
        }

        public static void Initialize() {
            // todo: mutex?
            pipeFactory = new DuplexChannelFactory<IServiceContract>(
                    new CommCallback(),
                    new NetNamedPipeBinding(NetNamedPipeSecurityMode.None),
                    new EndpointAddress("net.pipe://localhost/" + ServiceConst.PIPE_NAME));
            commChannel = pipeFactory.CreateChannel();
            try {
                commChannel.Subscribe(WindowsIdentity.GetCurrent().User.ToString());
            }
            catch(EndpointNotFoundException) {
                // todo: restart service
            }
        }

        public static void StaticBroadcast(Action action) {
            commChannel.Broadcast(QTUtility.ObjectToByteArray(new SerializeDelegate(action)));
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
            if(commChannel.ExecuteOnMainProcess(QTUtility.ObjectToByteArray(new SerializeDelegate(action)))) {
                action();
            }
        }

        public static void InvokeMain(Action<QTTabBarClass> action) {
            ExecuteOnMainProcess(() => LocalInvokeMain(action));
        }

        private static void LocalInvokeMain(Action<QTTabBarClass> action) {
            QTTabBarClass instance;
            using(new Keychain(rwLockTabBar, false)) {
                instance = mainInstance;
            }
            if(instance != null) action(instance);
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
            using(new Keychain(rwLockTabBar, true)) {
                dictTabInstances[Thread.CurrentThread] = tabbar;
                dictTabHandles[tabbar.Handle] = tabbar;
                mainInstance = tabbar;
            }
            commChannel.PushInstance(tabbar.Handle);
        }

        public static void UnregisterButtonBar() {
            using(new Keychain(rwLockBtnBar, true)) {
                dictBBarInstances.Remove(Thread.CurrentThread);
            }
        }

        public static bool UnregisterTabBar() {
            using(new Keychain(rwLockTabBar, true)) {
                QTTabBarClass tab;
                if(dictTabInstances.TryGetValue(Thread.CurrentThread, out tab)) {
                    dictTabInstances.Remove(Thread.CurrentThread);
                    dictTabHandles.Remove(tab.Handle);
                    // The MainInstance will be set by the callback, but just in case, set it here
                    // so that it's never set to an invalid value.
                    mainInstance = dictTabInstances.Count > 0 ? dictTabInstances.Values.First() : null;
                    commChannel.DeleteInstance(tab.Handle);
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
