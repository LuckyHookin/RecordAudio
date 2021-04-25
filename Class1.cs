using NAudio.Wave;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace RecordAudio
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual), IDispatchImpl(IDispatchImplType.InternalImpl)]
    class Record
    {
        private static WasapiLoopbackCapture capture = null;

        public static string startRecord(string path, int maxTimeSecond)
        {
            var outputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wave");
            if (path != null)
            {
                outputFolder = path;
            }
            if (maxTimeSecond == 0)
            {
                maxTimeSecond = 60;
            }
            //test
            Console.Out.WriteLine(outputFolder);
            Directory.CreateDirectory(outputFolder);
            var fileName = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss") + "-Record.wav";
            var outputFilePath = Path.Combine(outputFolder, fileName);
            capture = new WasapiLoopbackCapture();
            var writer = new WaveFileWriter(outputFilePath, capture.WaveFormat);
            capture.DataAvailable += (s, a) =>
            {
                writer.Write(a.Buffer, 0, a.BytesRecorded);
                if (writer.Position > capture.WaveFormat.AverageBytesPerSecond * maxTimeSecond)
                {
                    //test
                    Console.Out.WriteLine($"out Time:{maxTimeSecond}");
                    capture.StopRecording();
                }
            };
            capture.RecordingStopped += (s, a) =>
            {
                writer.Dispose();
                writer = null;
                capture.Dispose();
            };
            capture.StartRecording();
            return outputFilePath;
        }
        public static bool stopRecord()
        {
            if (capture == null)
            {
                return false;
            }
            if (capture.CaptureState != NAudio.CoreAudioApi.CaptureState.Stopped)
            {
                capture.StopRecording();
                return true;
            }

            return false;
        }
    }
}
