﻿using SharpDX.Direct3D11;
using SharpDX.DXGI;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestFrame.Model
{
    public class DirectXOutput : IDisposable
    {
        public DirectXOutput(Adapter1 adapter,
            SharpDX.Direct3D11.Device device,
            OutputDuplication outputDuplication,
            Texture2D texture2D,
            DisplayModeRotation rotation)
        {
            Adapter = adapter;
            Device = device;
            OutputDuplication = outputDuplication;
            Texture2D = texture2D;
            Rotation = rotation;
        }

        public Adapter1 Adapter { get; }
        public SharpDX.Direct3D11.Device Device { get; }
        public OutputDuplication OutputDuplication { get; }
        public DisplayModeRotation Rotation { get; }
        public Texture2D Texture2D { get; }

        public void Dispose()
        {
            OutputDuplication.ReleaseFrame();
            Disposer.TryDisposeAll(OutputDuplication, Texture2D, Adapter, Device);
            GC.SuppressFinalize(this);
        }
    }
}
