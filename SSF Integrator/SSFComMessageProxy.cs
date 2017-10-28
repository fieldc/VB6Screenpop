/****************************** Module Header ******************************\
* Module Name:  SSFComMessageProxy.cs
* Project:      SSFIntegration
* Copyright (c) Microsoft Corporation.
* 
* The definition of the COM class, SSFComMessageProxy, and its ClassFactory, 
* SSFComMessageProxyClassFactory.
* 
* (Please generate new GUIDs when you are writing your own COM server) 
* Program ID: SSFIntegration.SSFComMessageProxy
* CLSID_SimpleObject: DB9935C1-19C5-4ed2-ADD2-9A57E19F53A3
* IID_ISimpleObject: 941D219B-7601-4375-B68A-61E23A4C8425
* DIID_ISimpleObjectEvents: 014C067E-660D-4d20-9952-CD973CE50436
* 
* Properties:
* // With both get and set accessor methods
* float FloatProperty
* 
* Methods:
* // HelloWorld returns a string "HelloWorld"
* string HelloWorld();
* // GetProcessThreadID outputs the running process ID and thread ID
* void GetProcessThreadID(out uint processId, out uint threadId);
* 
* Events:
* // FloatPropertyChanging is fired before new value is set to the 
* // FloatProperty property. The Cancel parameter allows the client to cancel 
* // the change of FloatProperty.
* void FloatPropertyChanging(float NewValue, ref bool Cancel);
* 
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
* All other rights reserved.
* 
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#region Using directives
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;
#endregion


namespace SSFIntegration
{
    #region NYL Helper Classes
    internal class CallbackWorkItem
    {
        public string clientId;
        public string taxId;
        public string contractId;
        public int contractCount; 
        public IRecordRequestCallback callback;
        public CallbackWorkItem(IRecordRequestCallback cb, string clientId, string taxId, string contractId,int contractCount)
        {
            this.callback = cb;
            this.clientId = clientId;
            this.taxId = taxId;
            this.contractId = contractId;
            this.contractCount = contractCount;
        }
    }
    internal class BackendSingleton
    {
        public event SSFComMessageProxy.RequestedEventHandler _Requested;

        private static Guid id = Guid.NewGuid();
        public List<IRecordRequestCallback> callbacks;
        private List<IRecordRequestCallback> deadCallbacks;
        private static readonly BackendSingleton instance = new BackendSingleton();
        
        private BackendSingleton() {
            Debug.WriteLine("BackendSingleton Id=" + id.ToString()); 
            Debug.WriteLine("Created in pid=" + Process.GetCurrentProcess().Id);
            this.callbacks = new List<IRecordRequestCallback>();
            this.deadCallbacks = new List<IRecordRequestCallback>();
        }
        
        /// <summary>
        /// Our singleton instance pointer
        /// </summary>
        public static BackendSingleton Instance
        {
            get { return instance; }
        }

        /// <summary>
        /// Callback for threadpool worker
        /// </summary>
        /// <param name="callbackInfo"></param>
        public void doCallback(object callbackInfo)
        {
            CallbackWorkItem cwi = (CallbackWorkItem)callbackInfo;
            try
            {
                cwi.callback.Callback(cwi.clientId,cwi.taxId, cwi.contractId,cwi.contractCount);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                lock (this.deadCallbacks)
                {
                    this.deadCallbacks.Add(cwi.callback);
                }
            }
        }
        

        /// <summary>
        /// Raise the record request event by calling back the registered 
        /// handlers
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="contractId"></param>
        public void RaiseRecordRequested(string clientId, string taxId, string contractId,int contractCount)
        {
            if (this._Requested != null)
            {
                this._Requested(clientId, taxId, contractId, contractCount);
            }
            lock (this)
            {
                if (this.callbacks.Count>0)
                {
                    foreach (IRecordRequestCallback callback in this.callbacks)
                    {
                        if (this.deadCallbacks.Contains(callback))
                        {
                            Debug.WriteLine("Skipping deadcallbacks");
                            continue;
                        }

                        try
                        {
                            ThreadPool.QueueUserWorkItem(this.doCallback,new CallbackWorkItem(callback,clientId,taxId,contractId,contractCount));
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                    }
                    foreach (IRecordRequestCallback callback in this.deadCallbacks)
                    {
                        Debug.WriteLine("Removing Callback");
                        this.callbacks.Remove(callback);
                        
                    }
                    lock (this.deadCallbacks)
                    {
                        this.deadCallbacks.Clear();
                    }
                }
            }
        }

        /// <summary>
        /// Register a hander to be called back
        /// </summary>
        /// <param name="cb"></param>
        public void RegisterCallback(IRecordRequestCallback cb)
        {
            lock (this)
            {
                if (!this.callbacks.Contains(cb))
                    this.callbacks.Add(cb);
            }
        }
        
        /// <summary>
        /// Remove your registered callback
        /// </summary>
        /// <param name="cb"></param>
        public void UnRegisterCallback(IRecordRequestCallback cb)
        {
            lock (this)
            {
                if (this.callbacks.Contains(cb))
                    this.callbacks.Remove(cb);
            }
        }
        
        /// <summary>
        /// Clear all callbacks
        /// </summary>
        public void ClearCallbacks()
        {
            lock (this)
            {
                this.callbacks.Clear();
            }
        }
    }
    #endregion

    #region Interfaces


    /// <summary>
    /// This is the interface implemented by clients that wish 
    /// to be notified when the desktop is requesting an SSF record
    /// be pulled
    /// </summary>
    [Guid(SSFComMessageProxy.CallBackInterfaceId), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRecordRequestCallback
    {
        #region Properties

        #endregion
        #region Methods
        [DispId(1)]
        void Callback(string clientId, string taxId, string contractId, int contractCount);
        #endregion
    }

    [Guid(SSFComMessageProxy.EventsId), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRecordRequestedEvent
    {
        #region Properties

        #endregion
        #region Methods
        [DispId(1)]
        void Requested(string clientId, string taxId, string contractId, int contractCount);
        #endregion
    }

    /// <summary>
    /// This is the interface implemented by our out of process COM object
    /// to allow clients to register and raise callbacks between processes
    /// </summary>
    [Guid(SSFComMessageProxy.InterfaceId), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRecordRequest
    {
        #region Methods
        [DispId(1)]
        void RaiseRecordRequested(string clientId, string taxId, string contractId, int contractCount);
        [DispId(2)]
        void RegisterCallBack(IRecordRequestCallback cb);
        [DispId(3)]
        void UnRegisterCallBack(IRecordRequestCallback cb);
        #endregion
    }

    
    #endregion

    [ClassInterface(ClassInterfaceType.None)]           // No ClassInterface
    [ProgIdAttribute("SSFIntegration.SSFComMessageProxy")]
    [Guid(SSFComMessageProxy.ClassId), ComVisible(true)]
    [ComSourceInterfaces(typeof(IRecordRequestedEvent))]
    [ComDefaultInterface(typeof(IRecordRequest))]
    public class SSFComMessageProxy : ReferenceCountedObject, IRecordRequest
    {
        public delegate void RequestedEventHandler(string clientId, string taxId, string contractId, int contractCount);
        public event RequestedEventHandler Requested
        {
            add
            {
                BackendSingleton.Instance._Requested += value;
            }
            remove
            {
                BackendSingleton.Instance._Requested -= value;
            }
        }

        public const string ClassId = "8E15D03B-C5CD-497F-BA84-FA4CBBDC3CE7";
        public const string InterfaceId = "458467BD-FC0A-4E41-A0DB-846EAD737C05";
        public const string CallBackInterfaceId = "BB66D7D5-1AD4-45D5-B5F6-2C8F045D7234";
        public const string EventsId = "CE1D2C65-0BD8-4B34-9336-8B5FF39370E3";
        
        public SSFComMessageProxy() { }

        #region Methods
        public void RegisterCallBack(IRecordRequestCallback requestor)
        {
            BackendSingleton.Instance.RegisterCallback(requestor);
        }

        /// <summary>
        /// Register yourself to get a notification that there
        /// has been a request for a record
        /// </summary>
        /// <param name="remove"></param>
        public void UnRegisterCallBack(IRecordRequestCallback remove)
        {
            BackendSingleton.Instance.UnRegisterCallback(remove);
        }
        
        /// <summary>
        /// Notify SSF of record request
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="contractId"></param>
        public void RaiseRecordRequested(string clientId, string taxId, string contractId, int contractCount)
        {
            BackendSingleton.Instance.RaiseRecordRequested(clientId, taxId,contractId,contractCount);
        }
        #endregion

        #region COM Helper Functions
        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction()]
        public static void Register(Type t)
        {
            try
            {
                Debug.WriteLine("SSFComMessageProxy.Register()");
                COMHelper.RegasmRegisterLocalServer(t);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw ex; // Re-throw the exception
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction()]
        public static void Unregister(Type t)
        {
            try
            {
                Debug.WriteLine("SSFComMessageProxy.Unregister()");
                COMHelper.RegasmUnregisterLocalServer(t);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message); // Log the error
                throw ex; // Re-throw the exception
            }
        }

        #endregion
    }

    #region Class Factory
    /// <summary>
    /// Class factory for the class SSFComMessageProxy.
    /// </summary>
    internal class SSFComMessageProxyClassFactory : IClassFactory
    {
        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
        {
            ppvObject = IntPtr.Zero;

            if (pUnkOuter != IntPtr.Zero)
            {
                // The pUnkOuter parameter was non-NULL and the object does 
                // not support aggregation.
                Marshal.ThrowExceptionForHR(COMNative.CLASS_E_NOAGGREGATION);
            }

            if (riid == new Guid(SSFComMessageProxy.ClassId) ||
                riid == new Guid(COMNative.IID_IDispatch) ||
                riid == new Guid(COMNative.IID_IUnknown))
            {
                // Create the instance of the .NET object
                ppvObject = Marshal.GetComInterfaceForObject(
                    new SSFComMessageProxy(), typeof(IRecordRequest));
            }
            else
            {
                // The object that ppvObject points to does not support the 
                // interface identified by riid.
                Marshal.ThrowExceptionForHR(COMNative.E_NOINTERFACE);
            }

            return 0;   // S_OK
        }

        public int LockServer(bool fLock)
        {
            return 0;   // S_OK
        }
    }
    #endregion


    /// <summary>
    /// Reference counted object base.
    /// </summary>
    [ComVisible(false)]
    public class ReferenceCountedObject
    {
        public ReferenceCountedObject()
        {
            // Increment the lock count of objects in the COM server.
            ExeCOMServer.Instance.Lock();
        }

        ~ReferenceCountedObject()
        {
            // Decrement the lock count of objects in the COM server.
            ExeCOMServer.Instance.Unlock();
        }
    }
}
