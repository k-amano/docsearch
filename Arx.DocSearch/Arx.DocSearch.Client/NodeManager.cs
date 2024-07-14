using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Arx.DocSearch.Client
{
    public class NodeManager
    {
        public const string DllFileName = "NodeManager_Free.dll";
        public enum TNodeManagerKind
        {
            NMKNone, NMKClient, NMKAgent, NMKBoth
        }


        [DllImport(DllFileName)]
        public extern static void NMInitializeA(string kind);
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
        public extern static void NMStopProgram(long UserIndex);

        [DllImport(DllFileName)]
        public extern static void NMGetCluster(ref uint DResult);

        [DllImport(DllFileName)]
        public extern static void NMGetBoardCount(uint DCluster, ref long Result);
        [DllImport(DllFileName)]
        public extern static void NMGetBoard(uint DCluster, long BoardIndex, ref uint DResult);

    }
}
