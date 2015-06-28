using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Timers;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

// Based on Kenny Kerr's code: http://weblogs.asp.net/kennykerr/Rtd3 Thanks Kenny!
// NB this requires a Reference to the Microsoft.Office.Interop.Excel assembly to build.
// You also need to run regasm to register the DLL so that Excel's lookup
// of the ProgId parameter to the RTD( ) call will work. 
// You do *not* need to go to Excel's addin manager GUI to add this DLL. The COM registration
// takes care of allowing Excel to find the DLL location. For .Net 4.0 this cmd line
// does the registration...
// C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe kkaddin.dll /codebase
// It worked for me without admin permissions.
// Don't forget to set Formulas/Calculations Options to automatic to see the ticking cell!
// JOS 2015-06-28

namespace kkaddin {
    [
        Guid( "B6AF4673-200B-413c-8536-1F778AC14DE1" ),
        ProgId( "kkaddin.RtdServer" ),
        ComVisible( true )
    ]
    public class RtdServer : IRtdServer {
        private IRTDUpdateEvent m_callback;
        private Timer m_timer;
        private Dictionary<int, string> m_topics;

        public int ServerStart( IRTDUpdateEvent callback ) {
            m_callback = callback;
            m_timer = new Timer( );
            m_timer.Elapsed += this.OnTimerEvent;
            m_timer.Interval = 2000;
            m_topics = new Dictionary<int, string>( );
            return 1;
        }

        public void ServerTerminate( ) {
            if (null != m_timer) {
                m_timer.Dispose( );
                m_timer = null;
            }
        }

        public object ConnectData( int topicId,
                                  ref Array strings,
                                  ref bool newValues ) {
            if (1 != strings.Length) {
                return "Exactly one parameter is required (e.g. 'hh:mm:ss').";
            }
            string format = strings.GetValue( 0 ).ToString( );
            m_topics[topicId] = format;
            m_timer.Start( );
            return GetTime( format );
        }

        public void DisconnectData( int topicId ) {
            m_topics.Remove( topicId );
        }

        public Array RefreshData( ref int topicCount ) {
            object[,] data = new object[2, m_topics.Count];
            int index = 0;
            foreach (int topicId in m_topics.Keys) {
                data[0, index] = topicId;
                data[1, index] = GetTime( m_topics[topicId] );
                ++index;
            }
            topicCount = m_topics.Count;
            m_timer.Start( );
            return data;
        }

        public int Heartbeat( ) {
            return 1;
        }

        protected void OnTimerEvent( object o, ElapsedEventArgs e ) {
            m_timer.Stop( );
            m_callback.UpdateNotify( );
        }

        private static string GetTime( string format ) {
            return DateTime.Now.ToString( format, CultureInfo.CurrentCulture );
        }
    }
}
