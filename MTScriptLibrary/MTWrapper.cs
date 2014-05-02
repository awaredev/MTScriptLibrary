using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace MTScriptLibrary
{
   
    /// <summary>
    /// Base class for starting and controlling a Meditech Client session
    /// </summary>
    public class MtScriptSession 
    {
        //Timeout Variable, defaults to 0 (NO timeout)
        private uint _timeout;

        #region Properties
       
        /// <summary>
        /// Read Only property of the title of the current Meditech Window/Message Box
        /// </summary>
        public string Title
        {
            get
            {
                IntPtr buff = Marshal.AllocCoTaskMem(1);//Buffer to hold title
                IntPtr ptSz = Marshal.AllocCoTaskMem(4); //Buffer to hold size
                Marshal.WriteInt32(ptSz ,1);//Write base size of 1 to buffer
                string title; //String to contain title for return
                //Use Try to make sure errors don't leave unmanaged memory hooks
                try
                {
                    MtWrapper.MCSSetTimeout(1);//Timeout ensures DLL won't get caught in a loop
                    Int32 count = 0;//Count number of atempts
                    Int32 res = MtWrapper.MCSGetTitle(ptSz, buff);//Try to get title                 
                    //Console.WriteLine(res.ToString()); Troubleshooting line
                    while (res < 0)
                    {
                        if (res == -4) //Buffer too small, DLL will set ptSz mem location to correct size
                        {
                            Marshal.ReAllocCoTaskMem(buff, Marshal.ReadInt32(ptSz));//resize buffer
                            res = MtWrapper.MCSGetTitle(ptSz, buff); }//Try again
                        else if (res == -2)//Timeout- DLL will often timeout on first atempt, 
                        { 
                            res = MtWrapper.MCSGetTitle(ptSz, buff);
                            if (count < 2) //Increment count the first 2 timeouts
                            { count++; }
                            else
                            { throw new MeditechTimeoutException(this); }//After 3 timeouts,throw error
                        }
                        else if (res == -1)//Linked Client Window is no longer present
                        { throw new MeditechNoWindowException(this); }
                    }
                    //Read string from buffer pointer
                    title = Marshal.PtrToStringAnsi(buff, Marshal.ReadInt32(ptSz)); 
                }
                catch 
                { throw; }//No internal error handling for now
                finally
                {
                    //Free Pointers
                    Marshal.FreeCoTaskMem(buff);
                    Marshal.FreeCoTaskMem(ptSz);
                }
                MtWrapper.MCSSetTimeout(_timeout);//Return Timeout to default
                return title;//Return Window Title
            }
        }

        /// <summary>
        /// Property for getting/setting the timeout for Meditech commands. Set to 0 for no timeout (default)
        /// </summary>
        public uint Timeout
        {
            get
            {
                return _timeout;
            }
            set
            {
                _timeout = value; //Store locally before sending to script DLL
                Int32 res = MtWrapper.MCSSetTimeout(value);
                if (res == -1)
                { throw new MeditechNoWindowException(this);}
            }
        }

        /// <summary>
        /// Read Only property for accessing the current Meditech Window properties
        /// </summary>
        public MtWindow CurrentWindow
        {
            get 
            {               
                var win = new LpmcsInfo(); //Base object for managed code
                IntPtr winPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(win));//Allocate unmanaged memory for object
                Marshal.StructureToPtr(win, winPtr, true); //copy base object to unmanaged memory for DLL access
                Int32 count = 0; //Track number of atempts in case of loop
                MtWrapper.MCSSetTimeout(1); //Set timeout so DLL won't get caught in loop

                try
                {
                    int res = MtWrapper.MCSGetInfo(winPtr); //Call DLL to get Current Window
                    while (res < 0)
                    {
                        if (res == -1)//No Window error from DLL
                        { throw new MeditechNoWindowException(this); }
                        if (res == -2 && count < 2)//Loop if less than 3 atempts
                        {
                            count++; //Increment count and try again
                            res = MtWrapper.MCSGetInfo(winPtr);
                        }
                        else if (res == -2)// On third timeout, throw timeout exception
                        { throw new MeditechTimeoutException(this); }
                    }
                    Marshal.PtrToStructure(winPtr, win);//Copy DLL response from unmanaged Memory if response is OK (0)
                }
                catch { Console.WriteLine("error"); throw; }
                finally
                {
                    //Free Pointers
                    Marshal.FreeCoTaskMem(winPtr);
                }
                //Return a newly built friendly MT Window Object built from win object
                return new MtWindow(win);
            }
        }

        /// <summary>
        /// Read Only property displaying the contents of the current active Meditech field
        /// </summary>
        public string CurrentField
        {
            get 
            {                
                IntPtr buff = Marshal.AllocCoTaskMem(1); //Buffer to hold title
                IntPtr ptSz = Marshal.AllocCoTaskMem(4); //Buffer to hold size          
                Marshal.WriteInt32(ptSz, 1); //Set size to 1 
                string field; //String to contain field data for return
                //use Try to make sure errors don't leave unmanaged memory hooks
                try
                {
                    Int32 count = 0; //Counter to track atempts
                    Int32 res = MtWrapper.MCSGetCurrentData(ptSz, buff); //Call DLL
                    while (res < 0)
                    {
                        if (res == -4) //Buffer too small, szPtr will now point to correct size
                        {
                            Marshal.ReAllocCoTaskMem(buff, Marshal.ReadInt32(ptSz));//resize buffer
                            res = MtWrapper.MCSGetCurrentData(ptSz, buff); }//Try again
                        else if (res == -2 && count <2) //Timeout with less than 2 tries
                        { count++; //Increment count
                            res = MtWrapper.MCSGetCurrentData(ptSz, buff); }//Try again
                        else if (res == -2) //Too many atempts timed out, throw error
                        { throw new MeditechTimeoutException(this); }
                        else if (res == -1) //No Window
                        { throw new MeditechNoWindowException(this); }
                    }
                    //Read field string from pointer location
                    field = Marshal.PtrToStringAnsi(buff, Marshal.ReadInt32(ptSz)); 
                }
                catch//(Exception ex)
                { throw; }//No internal error handling for now
                finally
                {
                    //Free Pointers
                    Marshal.FreeCoTaskMem(buff);
                    Marshal.FreeCoTaskMem(ptSz);
                }
                //Return field string
                return field;
            }
        }
        #endregion

        #region Enumerations

        /// <summary>
        /// Enumeration of special key BYTE values for sending special keystrokes to the Meditech Client
        /// </summary>
        public enum SpecialKey
        {
            Escape = 27,
            Backspace = 8,
            Tab = 9,
            Delete = 127,
            Insert = 128,
            End = 129,
            Home = 130,
            PageDown = 131,
            PageUp = 132,
            UpArrow = 133,
            DownArrow = 134,
            LeftArrow = 135,
            RightArrow = 136,
            Return = 13,
            F1 = 137,
            F2 = 138,
            F3 = 139,
            F4 = 140,
            F5 = 141,
            F6 = 142,
            F7 = 143,
            F8 = 144,
            F9 = 145,
            F10 = 146,
            F11 = 147,
            F12 = 148,
            //SHIFT
            Shift_Backspace = 30,
            Shift_Tab = 11,
            Shift_Delete = 159,
            Shift_Insert = 160,
            Shift_End = 161,
            Shift_Home = 162,
            Shift_PageDown = 163,
            Shift_PageUp = 164,
            Shift_UpArrow = 165,
            Shift_DownArrow = 166,
            Shift_LeftArrow = 167,
            Shift_RightArrow = 168,
            Shift_F1 = 169,
            Shift_F2 = 170,
            Shift_F3 = 171,
            Shift_F4 = 172,
            Shift_F5 = 173,
            Shift_F6 = 174,
            Shift_F7 = 175,
            Shift_F8 = 176,
            Shift_F9 = 177,
            Shift_F10 = 178,
            Shift_F11 = 179,
            Shift_F12 = 180,
            //CTRL
            Ctrl_Backspace = 31,
            Ctrl_Tab = 28,
            Ctrl_Delete = 191,
            Ctrl_Insert = 192,
            Ctrl_End = 193,
            Ctrl_Home = 194,
            Ctrl_PageDown = 195,
            Ctrl_PageUp = 196,
            Ctrl_UpArrow = 197,
            Ctrl_DownArrow = 198,
            Ctrl_LeftArrow = 199,
            Ctrl_RightArrow = 200,
            Ctrl_F1 = 201,
            Ctrl_F2 = 202,
            Ctrl_F3 = 203,
            Ctrl_F4 = 204,
            Ctrl_F5 = 205,
            Ctrl_F6 = 206,
            Ctrl_F7 = 207,
            Ctrl_F8 = 208,
            Ctrl_F9 = 209,
            Ctrl_F10 = 210,
            Ctrl_F11 = 211,
            Ctrl_F12 = 212,
            //SHFT-CTRL
            ShftCtrl_Tab = 29,
            ShftCtrl_Delete = 223,
            ShftCtrl_Insert = 224,
            ShftCtrl_End = 225,
            ShftCtrl_Home = 226,
            ShftCtrl_PageDown = 227,
            ShftCtrl_PageUp = 228,
            ShftCtrl_UpArrow = 229,
            ShftCtrl_DownArrow = 230,
            ShftCtrl_LeftArrow = 231,
            ShftCtrl_RightArrow = 232,
            ShftCtrl_F1 = 233,
            ShftCtrl_F2 = 234,
            ShftCtrl_F3 = 235,
            ShftCtrl_F4 = 236,
            ShftCtrl_F5 = 237,
            ShftCtrl_F6 = 238,
            ShftCtrl_F7 = 239,
            ShftCtrl_F8 = 240,
            ShftCtrl_F9 = 241,
            ShftCtrl_F10 = 242,
            ShftCtrl_F11 = 243,
            ShftCtrl_F12 = 244,
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructor for MTScriptSession. Launches the client and provides access to the scripting commands for this instance
        /// </summary>
        /// <param name="pathToClient">Relative or Absolute (on same drive) path to the Meditech Client Application</param>
        /// <param name="minimized">boolean value as to whether or not to launch client minimized. 
        /// MT Scripting DLL cannot handle certain popups when maximized, so this should be set to TRUE</param>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        public MtScriptSession(string pathToClient, bool minimized)
        {
            //Check Path to see if it is an absolute path and convert to relative
            if (System.Text.RegularExpressions.Regex.IsMatch(pathToClient, "^[a-z|A-Z]:"))//Check for Drive Letter
            {
                pathToClient = System.Text.RegularExpressions.Regex.Replace(pathToClient, @"^[a-z|A-Z]:\\", "");//Remove letter
                string backstep = "";//String for leading directory up string
                int steps = Environment.CurrentDirectory.Split(char.Parse(@"\")).Length -1;//Determine steps to root
                for (int i =1; i <= steps; i++)//Write steps
                    {
                    backstep += @"..\";
                    }
                //Console.WriteLine(Environment.CurrentDirectory); //Troubleshooting
                pathToClient = backstep + pathToClient; //Converge strings
                pathToClient = pathToClient.Replace("\\", "/"); //Convert backslash to C friendly forwardslash
            }
            int min = 0; //Minimizer integer argument
            if (minimized)
            { min = 1; }
            int output = MtWrapper.MCSStartup(pathToClient, min);
            //Console.WriteLine(output + " " + pathToClient); //Troubleshooting line
            if (output == -1)
            { throw new MeditechNoWindowException(this); }
            if (output == -2)
            { throw new MeditechTimeoutException(this); }
            //An initial 'SendKeys' backspace request is sent to help prevent timeouts on new commands
            MtWrapper.MCSSetTimeout(1);
            SendSpecKey(SpecialKey.Backspace); //Sends backspace to initialize window, else timeout occurs on future requests
            MtWrapper.MCSSetTimeout(0);
        }

        #endregion

        #region Methods
        //Public method for closing the session
        /// <summary>
        /// Method for closing the Meditech Client Session
        /// </summary>
        /// <returns>Returns True if Meditech Window confirms close</returns>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        public bool Close()
        {
            int output = MtWrapper.MCSShutdown(); //DLL shutdown command
            if (output == -1)
            { throw new MeditechNoWindowException(this); }
            if (output == -2)
            { throw new MeditechTimeoutException(this); }

            return (output == 0);
        }

        /// <summary>
        /// Method for sending standard alphanumeric and symbol keystrokes to the meditech client 
        /// </summary>
        /// <param name="str">String to send to the Meditech Client as series of keystrokes</param>
        /// <param name="confirm">Boolean for whether to confirm field matches string sent before continuing execution</param>
        /// <returns>Returns True if string sent correctly</returns>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        public bool SendString(string str, bool confirm)
        {
            int resp; //Integer to check response
            IntPtr sendStr = Marshal.AllocCoTaskMem(str.Length); //Assign pointer
            try //Using a TRY block so that Marshalled memory gets released even on error
            {
                //Set timeout to 1 second to prevent loop
                MtWrapper.MCSSetTimeout(1);
                //Marshal writes the string as ANSI bytes and returns the unmanaged pointer of the data
                //String length tells the DLL how many keys/characters are in the string
                sendStr = Marshal.StringToHGlobalAnsi(str);
                resp = MtWrapper.MCSEnterKeys((uint)str.Length, sendStr);
                while (resp < 0)
                {
                    if (resp == -1) //Throw meaningful exceptions
                    { throw new MeditechNoWindowException(this); }
                    //else if (resp == -2 && count < 2) //Timeout with atempts left
                    //{ resp = MTWrapper.MCSEnterKeys((uint)str.Length, sendStr); }
                    if (resp == -2) //Too many timeouts
                    { throw new MeditechTimeoutException(this); }
                }
                
                if (confirm)
                {
                    while (CurrentField != str)
                    { 
                        //If the field does not match, make sure that typing is correct so far
                        if (CurrentField.Length > 0 && !str.StartsWith(CurrentField))
                        { throw new MeditechScriptError(this); }
                    }
                }
                
            }
            catch
            {
                //Error Handling to go here as needed?
                throw; 
            }
            finally
            {
                Marshal.FreeCoTaskMem(sendStr); //Release the memory!
                MtWrapper.MCSSetTimeout(_timeout); //Reset Timeout to default or manual setting 
            }
            return (resp == 0);
         }

        //Public function to send a password string as keystrokes
        /// <summary>
        /// Used to send a password to the Meditech client in the correct format required by Meditech
        /// </summary>
        /// <param name="str">The Password as a string</param>
        /// <returns>Returns True if password inputed correctly</returns>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        public bool SendPassword(string str)
        {
            int resp = -1; //Response set to -1 until DLL call
            int count = 0; //Count for timed out atempts
            IntPtr sendStr = new IntPtr();
            try //Using a TRY block so that Marshalled memory gets released even on error
            {
                //Set timeout to 1 second to allow keysend response
                MtWrapper.MCSSetTimeout(1);
                
                //Marshal writes the string as ANSI bytes and returns the unmanaged pointer of the data
                //String length tells the DLL how many keys/characters are in the string
                sendStr = Marshal.StringToHGlobalAnsi(str);
                resp = MtWrapper.MCSEnterPasswordKeys((uint)str.Length, sendStr);
                while (resp < 0)
                {
                    if (resp == -1) //Throw meaningful exceptions
                    { throw new MeditechNoWindowException(this); }
                    if (resp == -2 && count < 2) //Timeout with atempts left
                    { resp = MtWrapper.MCSEnterPasswordKeys((uint)str.Length, sendStr); }
                    else if (resp == -2) //Too many timeouts
                    { throw new MeditechTimeoutException(this); }
                }
                
                if (resp == -1) //Throw meaningful exceptions
                { throw new MeditechNoWindowException(this); }
                else if (resp == -2)
                { throw new MeditechTimeoutException(this); }
            }
            catch
            {
                //Error Handling to go here as needed?
                throw;
            }
            finally
            {
                Marshal.FreeCoTaskMem(sendStr); //Release the memory!
                MtWrapper.MCSSetTimeout(_timeout); //Reset Timeout to default or manual setting
            }
            return (resp == 0);
        }
        
        /// <summary>
        /// Method to send the byte code for Special Keystrokes (Esc, Return Function keys, etc.) to the Meditech Client
        /// </summary>
        /// <param name="sKey">Special Key Enumeration</param>
        /// <returns>Returns True if key inputed correctly</returns>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        public bool SendSpecKey(SpecialKey sKey)
        {
            int resp = -1; //Sets response to -1 by default
            IntPtr buff = Marshal.AllocHGlobal(1);
            try //Using a TRY block so that Marshalled memory gets released even on error
            {
                //Set timeout to 1 second to allow keysend response
                MtWrapper.MCSSetTimeout(1);
                //Marshal writes the value for the special key to buff pointer
                //buff alsways has a length of 1, sent to the MT Client
                Marshal.WriteByte(buff, (byte)sKey);
                resp = MtWrapper.MCSEnterKeys(1, buff);

                if (resp == -1) //Throw meaningful exceptions
                { throw new MeditechNoWindowException(this); }
    
            }
            catch
            {
                //Error Handling to go here as needed?
                throw;
            }
            finally
            {
                Marshal.FreeHGlobal(buff); //Release the memory!
                MtWrapper.MCSSetTimeout(_timeout); //Reset Timeout to default or manual setting
            }
            return (resp == 0);
        }

        /// <summary>
        /// Method for Selecting a Tab/Page in a Meditech Window (for Meditech Windows that contain Tabs)
        /// </summary>
        /// <param name="tabIndex">Index # of the Tab. Use CurrentWindow Property to find Tab information</param>
        /// <returns>Returns True if tab selected correctly</returns>
        /// <exception cref="MeditechNoWindowException">Throws MeditechNoWindowException if no window found</exception>
        /// <exception cref="MeditechTimeoutException">Throws MeditechTimeoutException if timeout exceeded</exception>
        /// <exception cref="MeditechInvalidParameterException">Throws MeditechInvalidParameterException if tab index is invalid</exception>
        public bool SelectTab(uint tabIndex)
        {
            int res = MtWrapper.MCSSelect(tabIndex);
            //error handling
            if (res == -5) //Invalid tab
                {
                     throw new MeditechInvalidParameterException(this); 
                }
            if (res == -2)//Throw meaningful exceptions
                {throw new MeditechTimeoutException(this);}
            if (res == -1) 
            { throw new MeditechNoWindowException(this); }
            return (res == 0);
        }

#endregion

        #region Support Structs and Classes
        /// <summary>
        /// Structure for accessing information concerning the Current Meditech Window
        /// </summary>
        public struct MtWindow
        {
            private readonly bool _isMsgBox; //Boolean for is a message box
            private readonly Int32 _x; // Column
            private readonly Int32 _y; // Row
            private readonly Int32 _nPages; // How many pages are there?
            private readonly Int32 _nCurPage; // What page are we on now?

            //Constructor
            /// <summary>
            /// Constructor for MTWindow object to contain details of current Meditech Window. Should not be used 
            /// by coders, only by the MTScriptSession instance
            /// </summary>
            /// <param name="raw">Internal structure from unmanaged C dll for accessing client</param>
            internal MtWindow(LpmcsInfo raw)
            {
                if (raw.x < 0)//This is a Message Box
                {
                    _isMsgBox = true;
                    _x = 0;
                    _y = 0;
                    _nPages = 0;
                    _nCurPage = 0;
                }
                else
                {
                    _isMsgBox = false;
                    _x = raw.x;
                    _y = raw.y;
                    _nPages = raw.nPages;
                    _nCurPage = raw.nPages;
                }
            }

            //Properties
            /// <summary>
            /// Boolean property to identify if WIndow is a Message Box.
            /// </summary>
            public bool IsMessageBox //Quick property to identify message box 
            {
                get
                {
                    try
                    {
                        return _isMsgBox;
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); return _isMsgBox; }
                }
            }
            /// <summary>
            /// Read Only Property giving current Character X position in Meditech Window (0 if Mesage Box)
            /// </summary>
            public Int32 XPosition //Char X position of current field
            {
                get
                {
                    try { return _x; }
                    catch (Exception ex){ Console.WriteLine(ex.Message); return _x; }
                }
            }
            /// <summary>
            /// Read Only Property giving current Character Y position in Meditech Window (0 if Mesage Box)
            /// </summary>
            public Int32 YPosition //Char Y position of current field
            {
                get
                {
                    try { return _y; }
                    catch (Exception ex) { Console.WriteLine(ex.Message); return _y; }
                }
            }
            /// <summary>
            /// Read Only property displaying # of Tabs or 'Pages' in current Meditech Window
            /// </summary>
            public Int32 TabCount // # of tabs-pages in this window
            {
                get
                {
                    try { return _nPages ; }
                    catch (Exception ex) { Console.WriteLine(ex.Message); return _nPages; }
                }
            }
            /// <summary>
            /// Read Only Property showing current active Tab or 'Page' in current Meditech Window
            /// </summary>
            public Int32 TabIndex // Current tab-page
            {
                get
                {
                    try { return _nCurPage; }
                    catch (Exception ex) { Console.WriteLine(ex.Message); return _nCurPage; }
                }
            }

        }

        /// <summary>
        /// Internal Struct used for reading Window data from MCSCRIPT DLL
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        internal struct LpmcsInfo
        {
            public Int32 x; // Column
            public Int32 y; // Row
            public Int32 nPages; // How many pages are there?
            public Int32 nCurPage; // What page is currently active in the window?
        }

        //Static Marshalling-Management class for sending commands to MT CS via the scripting DLL provided by Meditech
        //Internal, no XML documentation required
        internal static class MtWrapper
        {
            //MCSStartup opens the MT Client and links it to the Scripting session
            //Returns 0 = OK -1 = FAIL
            [DllImport("MCSCRIPT.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            public static extern Int32 MCSStartup(string lpszProgram, int bMinimize);

            //MCSGetTitle gets the MT Window name
            // 1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout, -4 FAIL Buffer too small
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSGetTitle(IntPtr iBufferSize, IntPtr ptrBuffer);

            //MCSGetInfo returns info on the current MT Window
            //1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout
            //pmsci parameter is a structure to pass the window information to
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSGetInfo(IntPtr pmcsi);
            
            //MCSGetCurrentData retrieves data in field, OR text in MessageBox
            // 1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout, -4 FAIL Buffer too small
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSGetCurrentData(IntPtr iBufferSize, IntPtr ptrBuffer);

            //MCSEnterKeys enters byte codes for keystrokes in the client
            //1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout
            //iBufferSize = # of keys, ptrBuffer points to buffer of 1 byte keycodes
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSEnterKeys(uint iBufferSize, IntPtr ptrBuffer);

            //MCSEnterPasswordKeys enters byte codes for keystrokes in the client FOR PASSWORD ENTRY
            //1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout
            //iBufferSize = # of keys, ptrBuffer points to buffer of 1 byte keycodes
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSEnterPasswordKeys(uint iBufferSize, IntPtr ptrBuffer);

            //MCSSelect mouse-clicks a TAB on theh MT Window (1 indexed tab array)
            //Returns OK if no tabs defined on Window, ie, a message box
            //1= Message BOx, 0=OK, -1= FAIL No Window, -2 FAIL Timeout, -5 = FAIL Invalid parameter
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSSelect(uint itemId);

            //MCSSetTimeout Sets the timeout for script operations
            // 0 = OK, -1 = FAIL No Window
            // Set to 0 (default) to disable timeout
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSSetTimeout(uint nSeconds);

            //MCSShutdown closes the session
            [DllImport("MCSCRIPT.dll")]
            public static extern Int32 MCSShutdown();

        }

        #endregion

    }

    //EXCEPTION CLASSES
    #region Exception Classes

    /// <summary>
    /// Basic Meditech Exception Type
    /// </summary>
    public class MeditechException : ApplicationException
    {
        public MeditechException(MtScriptSession session, bool getData, string baseMessage)
            : base(baseMessage + ProcessData(session, getData))
        { }

        private static string ProcessData(MtScriptSession session, bool getData)
        {
            //Gather session information
            if (getData)
            {
                string title, msgBox, pages, currPage, field;//Declare strings
                //Try to pull data from session
                try
                { title = session.Title; }
                catch { title = "UNAVAILABLE"; }
                try
                { msgBox = session.CurrentWindow.IsMessageBox.ToString(); }
                catch { msgBox = "UNAVAILABLE"; }
                try
                { pages = session.CurrentWindow.TabCount.ToString(CultureInfo.InvariantCulture); }
                catch { pages = "UNAVAILABLE"; }
                try
                { currPage = session.CurrentWindow.TabIndex.ToString(CultureInfo.InvariantCulture); }
                catch { currPage = "UNAVAILABLE"; }
                try
                { field = session.CurrentField; }
                catch { field = "UNAVAILABLE"; }
                //Set base message with data
                return "Title: " + title + Environment.NewLine + "Message Box: " + msgBox + Environment.NewLine
                    + "# Pages/Tabs: " + pages + Environment.NewLine + "Active Page/Tab: " + currPage + Environment.NewLine
                    + "Field Content: " + field;
            }
            return "";

        }
    }

    /// <summary>
    /// Exception thrown when no Meditech Window was found
    /// </summary>
    public class MeditechNoWindowException : MeditechException
    {
        public MeditechNoWindowException(MtScriptSession session)
            : base(session, false, "Meditech Window not found")
        { }
    }

    /// <summary>
    /// Exception thrown if the timeout for receiving a response from the Meditech Window expires
    /// </summary>
    public class MeditechTimeoutException : MeditechException
    {
        public MeditechTimeoutException(MtScriptSession session)
            : base(session, true, "Meditech Scripting response Timeout exceeded\n")
        { }
    }

    /// <summary>
    /// Exception thrown when a command sent to the Meditech Window has an unexpected result
    /// </summary>
    public class MeditechScriptError : MeditechException
    {
        public MeditechScriptError(MtScriptSession session)
            : base(session, true, "Unexpected result from Meditech Script Command\n")
        { }
    }

    /// <summary>
    /// Exception thrown if a command sent to a Meditech Window supplied an invalid parameter
    /// </summary>
    public class MeditechInvalidParameterException : MeditechException
    {
        public MeditechInvalidParameterException(MtScriptSession session)
            : base(session, true, "Invalid parameter for current Meditech Window\n")
        { }
    }

#endregion
}
