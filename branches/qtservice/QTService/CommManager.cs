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
                    new NetNamedPipeBinding(NetNamedPipeSecurityMode.None) { ReceiveTimeout = TimeSpan.MaxValue },
                    ServiceConst.PIPE_NAME);
            serviceHost.Open();
        }

        [ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Multiple, InstanceContextMode = InstanceContextMode.PerCall)]
        private class CommService : IServiceContract {
            private class User {
                public List<ICallbackContract> Callbacks = new List<ICallbackContract>();
                public StackDictionary<IntPtr, ICallbackContract> Instances = new StackDictionary<IntPtr, ICallbackContract>();
            }

            private static Dictionary<string, User> UserIDs = new Dictionary<string, User>();
            private static Dictionary<ICallbackContract, User> Users = new Dictionary<ICallbackContract, User>();

            private static void CheckConnections() {
                Predicate<ICallbackContract> pred = callback =>
                        ((ICommunicationObject)callback).State != CommunicationState.Opened;
                foreach(var user in Users.Values) {
                    user.Callbacks.RemoveAll(pred);
                    user.Instances.RemoveAllValues(c => !user.Callbacks.Contains(c));                    
                }
                foreach(var callback in Users.Keys.Where(c => pred(c)).ToList()) {
                    Users.Remove(callback);
                }
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

            private static User GetUser() {
                CheckConnections();
                ICallbackContract callback = GetCallback();
                User user = callback == null ? new User() : Users[callback];
                return user;
            }

            public bool ExecuteOnMainProcess(byte[] encodedAction) {
                ICallbackContract target;
                lock(Users) {
                    User user = GetUser();
                    if(IsMainProcess()) {
                        return true;
                    }
                    else {
                        if(user.Instances.Count == 0) return false;
                        target = user.Instances.Peek();
                    }
                }
                target.Execute(encodedAction);
                return false;
            }

            public void Broadcast(byte[] encodedAction) {
                ICallbackContract sender = GetCallback();
                Action async = () => {
                    List<ICallbackContract> targets;
                    lock(Users) {
                        User user = GetUser();
                        targets = user.Callbacks.Where(c => c != sender).ToList();
                    }
                    targets.ForEach(c => c.Execute(encodedAction));
                };
                async.BeginInvoke(null, null);
            }

            public void DeleteInstance(IntPtr hwnd) {
                lock(Users) {
                    User user = GetUser();
                    IntPtr main = IntPtr.Zero;
                    if(user.Instances.Count > 0) user.Instances.Peek(out main);
                    if(!user.Instances.Remove(hwnd) || hwnd != main) return;
                    while(user.Instances.Count > 0 && !user.Instances.Peek(out main).SetMain(main)) {
                        user.Instances.Pop();
                    }
                }
            }

            public bool IsMainProcess() {
                lock(Users) {
                    User user = GetUser();
                    return user.Instances.Count > 0 && GetCallback() == user.Instances.Peek();
                }
            }

            public void Subscribe(string userid) {
                lock(Users) {
                    try {
                        CheckConnections();
                        ICallbackContract callback = GetCallback();
                        User user;
                        if(!UserIDs.TryGetValue(userid, out user)) {
                            user = new User();
                            UserIDs[userid] = user;
                        }
                        Users[callback] = user;
                        if(!user.Callbacks.Contains(callback)) {
                            user.Callbacks.Add(callback);
                        }
                    }
                    catch {
                    }                    
                }
            }

            public void PushInstance(IntPtr hwnd) {
                lock(Users) {
                    User user = GetUser();
                    if(!user.Callbacks.Contains(GetCallback())) return; // hmmm....
                    user.Instances.Push(hwnd, GetCallback());
                }
            }

            public int GetTotalInstanceCount() {
                lock(Users) {
                    return GetUser().Instances.Count;
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
        void Subscribe(string userid);

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

        [OperationContract]
        int GetTotalInstanceCount();
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
            var removeMe = lstKeys.Where(s => match(dictionary[s])).ToList();
            foreach(var s in removeMe) {
                lstKeys.Remove(s);
                dictionary.Remove(s);
            }
            return removeMe.Count;
        }

        public bool TryGetValue(S key, out T value) {
            return dictionary.TryGetValue(key, out value);
        }

        public int Count { get { return lstKeys.Count; } }

        public ICollection<S> Keys { get { return dictionary.Keys; } }

        public ICollection<T> Values { get { return dictionary.Values; } }
    }
}
