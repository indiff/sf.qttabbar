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
using System.ServiceModel;

namespace QTTabBarService {
    internal class CommManager : IDisposable {
        private ServiceHost serviceHost;
        

        public CommManager() {
            serviceHost = new ServiceHost(
                    typeof(CommService),
                    new Uri[] { new Uri("net.pipe://localhost") });
            serviceHost.AddServiceEndpoint(
                    typeof(IServiceContract),
                    new NetNamedPipeBinding(NetNamedPipeSecurityMode.None),
                    ServiceConst.PIPE_NAME);
            serviceHost.Open();
        }

        [ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Reentrant, InstanceContextMode = InstanceContextMode.PerCall)]
        private class CommService : IServiceContract {
            private static List<ICallbackContract> m_Callbacks = new List<ICallbackContract>();
            private static StackDictionary<IntPtr, ICallbackContract> m_Instances = new StackDictionary<IntPtr, ICallbackContract>();

            private static void CheckConnections() {
                m_Callbacks.RemoveAll(callback =>
                        ((ICommunicationObject)callback).State != CommunicationState.Opened);
                m_Instances.RemoveAllValues(c => !m_Callbacks.Contains(c));
            }

            private static ICallbackContract GetCallback() {
                ICallbackContract callback = null;
                try {
                    callback = OperationContext.Current.GetCallbackChannel<ICallbackContract>();
                }
                catch {
                }
                return callback;
            }

            public bool ExecuteOnMainProcess(byte[] encodedAction) {
                lock(m_Callbacks) {
                    CheckConnections();
                    if(IsMainProcess()) {
                        return true;
                    }
                    else {
                        if(m_Instances.Count > 0) {
                            m_Instances.Peek().Execute(encodedAction);
                        }
                        return false;
                    }
                }
            }

            public void Broadcast(byte[] encodedAction) {
                ICallbackContract sender = GetCallback();
                Action async = () => {
                    lock(m_Callbacks) {
                        CheckConnections();
                        foreach(ICallbackContract callback in m_Callbacks) {
                            if(callback != sender) {
                                callback.Execute(encodedAction);
                            }
                        }
                    }
                };
                async.BeginInvoke(null, null);
            }

            public void DeleteInstance(IntPtr hwnd) {
                lock(m_Callbacks) {
                    CheckConnections();
                    IntPtr main = IntPtr.Zero;
                    if(m_Instances.Count > 0) m_Instances.Peek(out main);
                    if(!m_Instances.Remove(hwnd) || hwnd != main) return;
                    while(m_Instances.Count > 0 && !m_Instances.Peek(out main).SetMain(main)) {
                        m_Instances.Pop();
                    }
                }
            }

            public bool IsMainProcess() {
                lock(m_Callbacks) {
                    CheckConnections();
                    return m_Instances.Count > 0 && GetCallback() == m_Instances.Peek();
                }
            }

            public void Subscribe() {
                lock(m_Callbacks) {
                    try {
                        ICallbackContract callback = OperationContext.Current.GetCallbackChannel<ICallbackContract>();
                        if(!m_Callbacks.Contains(callback)) {
                            m_Callbacks.Add(callback);
                        }
                    }
                    catch {
                    }                    
                }
            }

            public void PushInstance(IntPtr hwnd) {
                lock(m_Callbacks) {
                    CheckConnections();
                    if(!m_Callbacks.Contains(GetCallback())) return; // hmmm....
                    m_Instances.Push(hwnd, GetCallback());
                }
            }
        }

        public void Dispose() {
            if(serviceHost != null) {
                serviceHost.Close();
                serviceHost = null;
            }
        }
    }

    [ServiceContract(SessionMode = SessionMode.Required, CallbackContract = typeof(ICallbackContract))]
    public interface IServiceContract {
        [OperationContract]
        void Subscribe();

        [OperationContract]
        void PushInstance(IntPtr hwnd);

        [OperationContract]
        void DeleteInstance(IntPtr hwnd);

        [OperationContract]
        bool IsMainProcess();

        [OperationContract]
        bool ExecuteOnMainProcess(byte[] encodedAction);

        [OperationContract]
        void Broadcast(byte[] encodedAction);
    }

    public interface ICallbackContract {
        [OperationContract]
        bool SetMain(IntPtr hwnd);

        [OperationContract]
        void Execute(byte[] encodedAction);
    }

    internal sealed class StackDictionary<S, T> {
        private Dictionary<S, T> dictionary;
        private List<S> lstKeys;

        public StackDictionary() {
            lstKeys = new List<S>();
            dictionary = new Dictionary<S, T>();
        }

        public T Peek() {
            S local;
            return popPeekInternal(false, out local);
        }

        public T Peek(out S key) {
            return popPeekInternal(false, out key);
        }

        public T Pop() {
            S local;
            return popPeekInternal(true, out local);
        }

        public T Pop(out S key) {
            return popPeekInternal(true, out key);
        }

        private T popPeekInternal(bool fPop, out S lastKey) {
            if(lstKeys.Count == 0) {
                throw new InvalidOperationException("This StackDictionary is empty.");
            }
            lastKey = lstKeys[lstKeys.Count - 1];
            T local = dictionary[lastKey];
            if(fPop) {
                lstKeys.RemoveAt(lstKeys.Count - 1);
                dictionary.Remove(lastKey);
            }
            return local;
        }

        public void Push(S key, T value) {
            lstKeys.Remove(key);
            lstKeys.Add(key);
            dictionary[key] = value;
        }

        public bool Remove(S key) {
            lstKeys.Remove(key);
            return dictionary.Remove(key);
        }

        public int RemoveAllValues(Predicate<T> match) {
            int ret = lstKeys.RemoveAll(s => match(dictionary[s]));
            if(ret > 0) {
                dictionary = lstKeys.ToDictionary(s => s, s => dictionary[s]);    
            }
            return ret;
        }

        public bool TryGetValue(S key, out T value) {
            return dictionary.TryGetValue(key, out value);
        }

        public int Count { get { return lstKeys.Count; } }

        public ICollection<S> Keys { get { return dictionary.Keys; } }

        public ICollection<T> Values { get { return dictionary.Values; } }
    }
}
