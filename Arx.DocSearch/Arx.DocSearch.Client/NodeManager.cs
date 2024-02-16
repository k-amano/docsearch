using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Arx.DocSearch.Client
{
    public class NodeManager
    {
        public const string DllFileName = "NodeManager.dll";
        public enum TNodeManagerKind
        {
            NMKNone, NMKClient, NMKAgent, NMKBoth
        }


        [DllImport(DllFileName)]
        public extern static void NMInitialize(string DLLFileName, TNodeManagerKind Kind);
        [DllImport(DllFileName)]
        public extern static void NMFinalize();
        [DllImport(DllFileName)]
        public extern static void NMLogIn(uint Board);
        [DllImport(DllFileName)]
        public extern static void NMLogOut();
        [DllImport(DllFileName)]
        public extern static void NMOpenConfig(uint OnChangeConfig);

        [DllImport(DllFileName)]
        public extern static void NMCloseConfig();

        [DllImport(DllFileName)]
        public extern static void NMStartProgram(long UserIndex, string DLLFileName, string Params,  uint ProcessHandle);

        [DllImport(DllFileName)]
        public extern static void NMStopProgram(long UserIndexe);

        [DllImport(DllFileName)]
        public extern static uint NMCluster(uint DResult);

        [DllImport(DllFileName)]
        public extern static long NMBoardCount(uint DCluster);
        [DllImport(DllFileName)]
        public extern static uint NMBoard(uint DCluster, long BoardIndex);

    }
}
