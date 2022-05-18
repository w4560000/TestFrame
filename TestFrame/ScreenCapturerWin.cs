using Microsoft.Win32;
using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using TestFrame.Model;

namespace TestFrame
{
    public class ScreenCapturerWin
    {
        private readonly Dictionary<string, int> _bitBltScreens = new Dictionary<string, int>();
        private readonly Dictionary<string, DirectXOutput> _directxScreens = new Dictionary<string, DirectXOutput>();
        private readonly object _screenBoundsLock = new object();

        public ScreenCapturerWin()
        {
            Init();
            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
        }

        public event EventHandler<Rectangle> ScreenChanged;

        private enum GetDirectXFrameResult
        {
            Success,
            Failure,
            Timeout,
        }

        public bool CaptureFullscreen { get; set; } = true;
        public bool NeedsInit { get; set; } = true;
        public string SelectedScreen { get; private set; } = Screen.PrimaryScreen.DeviceName;
        public void Dispose()
        {
            try
            {
                SystemEvents.DisplaySettingsChanged -= SystemEvents_DisplaySettingsChanged;
                ClearDirectXOutputs();
                GC.SuppressFinalize(this);
            }
            catch { }
        }
        public Rectangle CurrentScreenBounds { get; private set; } = Screen.PrimaryScreen.Bounds;

        public IEnumerable<string> GetDisplayNames() => Screen.AllScreens.Select(x => x.DeviceName);

        public Bitmap GetNextFrame()
        {
            lock (_screenBoundsLock)
            {
                Bitmap returnFrame = null;
                var frameCompletedEvent = new ManualResetEventSlim();

                // This is necessary to ensure SwitchToInputDesktop works.  Threads
                // that have hooks in the current desktop will not succeed.
                var captureThread = new Thread(() =>
                {
                    try
                    {
                        Win32Interop.SwitchToInputDesktop();

                        if (NeedsInit)
                        {
                            Logger.Write("Init needed in GetNextFrame.");
                            Init();
                            NeedsInit = false;
                        }

                        // Sometimes DX will result in a timeout, even when there are changes
                        // on the screen.  I've observed this when a laptop lid is closed, or
                        // on some machines that aren't connected to a monitor.  This will
                        // have it fall back to BitBlt in those cases.
                        // TODO: Make DX capture work with changed screen orientation.
                        if (_directxScreens.TryGetValue(SelectedScreen, out var dxDisplay) &&
                            dxDisplay.Rotation == DisplayModeRotation.Identity)
                        {
                            var (result, frame) = GetDirectXFrame();

                            if (result == GetDirectXFrameResult.Timeout)
                            {
                                return;
                            }

                            if (result == GetDirectXFrameResult.Success)
                            {
                                returnFrame = frame;
                                return;
                            }
                        }

                        returnFrame = GetBitBltFrame();

                    }
                    catch (Exception e)
                    {
                        Logger.Write(e);
                        NeedsInit = true;
                    }
                    finally
                    {
                        frameCompletedEvent.Set();
                    }
                });

                captureThread.SetApartmentState(ApartmentState.STA);
                captureThread.Start();

                frameCompletedEvent.Wait();

                return returnFrame;
            }

        }

        public int GetScreenCount() => Screen.AllScreens.Length;

        public int GetSelectedScreenIndex()
        {
            if (_bitBltScreens.TryGetValue(SelectedScreen, out var index))
            {
                return index;
            }
            return 0;
        }

        public Rectangle GetVirtualScreenBounds() => SystemInformation.VirtualScreen;

        public void Init()
        {
            Win32Interop.SwitchToInputDesktop();

            CaptureFullscreen = true;
            InitBitBlt();
            InitDirectX();

            ScreenChanged?.Invoke(this, CurrentScreenBounds);
        }

        public void SetSelectedScreen(string displayName)
        {
            lock (_screenBoundsLock)
            {
                if (displayName == SelectedScreen)
                {
                    return;
                }

                if (_bitBltScreens.ContainsKey(displayName))
                {
                    SelectedScreen = displayName;
                }
                else
                {
                    SelectedScreen = _bitBltScreens.Keys.First();
                }
                RefreshCurrentScreenBounds();
            }
        }

        private void ClearDirectXOutputs()
        {
            foreach (var screen in _directxScreens.Values)
            {
                try
                {
                    screen.Dispose();
                }
                catch { }
            }
            _directxScreens.Clear();
        }

        private Bitmap GetBitBltFrame()
        {
            try
            {
                var currentFrame = new Bitmap(CurrentScreenBounds.Width, CurrentScreenBounds.Height, PixelFormat.Format32bppArgb);

                using (var graphic = Graphics.FromImage(currentFrame))
                {
                    graphic.CopyFromScreen(CurrentScreenBounds.Left, CurrentScreenBounds.Top, 0, 0, new Size(CurrentScreenBounds.Width, CurrentScreenBounds.Height));
                }

                currentFrame = getnew(Image.FromHbitmap(currentFrame.GetHbitmap()), 0.5);

                return currentFrame;
            }
            catch (Exception ex)
            {
                Logger.Write(ex);
                Logger.Write("Capturer error in BitBltCapture.");
                NeedsInit = true;
            }

            return null;
        }
        private (GetDirectXFrameResult result, Bitmap frame) GetDirectXFrame()
        {
            try
            {
                var duplicatedOutput = _directxScreens[SelectedScreen].OutputDuplication;
                var device = _directxScreens[SelectedScreen].Device;
                var texture2D = _directxScreens[SelectedScreen].Texture2D;

                // Try to get duplicated frame within given time is ms
                var result = duplicatedOutput.TryAcquireNextFrame(500,
                    out var duplicateFrameInformation,
                    out var screenResource);

                if (result.Failure)
                {
                    if (result.Code == SharpDX.DXGI.ResultCode.WaitTimeout.Code)
                    {
                        return (GetDirectXFrameResult.Timeout, null);
                    }
                    else
                    {
                        Logger.Write($"TryAcquireFrame error.  Code: {result.Code}");
                        NeedsInit = true;
                        return (GetDirectXFrameResult.Failure, null);
                    }
                }

                if (duplicateFrameInformation.AccumulatedFrames == 0)
                {
                    try
                    {
                        duplicatedOutput.ReleaseFrame();
                    }
                    catch { }
                    return (GetDirectXFrameResult.Failure, null);
                }

                var currentFrame = new Bitmap(texture2D.Description.Width, texture2D.Description.Height, PixelFormat.Format32bppArgb);

                // Copy resource into memory that can be accessed by the CPU
                using (var screenTexture2D = screenResource.QueryInterface<Texture2D>())
                {
                    device.ImmediateContext.CopyResource(screenTexture2D, texture2D);
                }

                // Get the desktop capture texture
                var mapSource = device.ImmediateContext.MapSubresource(texture2D, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);

                var boundsRect = new Rectangle(0, 0, texture2D.Description.Width, texture2D.Description.Height);

                // Copy pixels from screen capture Texture to GDI bitmap
                var mapDest = currentFrame.LockBits(boundsRect, ImageLockMode.WriteOnly, currentFrame.PixelFormat);
                var sourcePtr = mapSource.DataPointer;
                var destPtr = mapDest.Scan0;
                for (int y = 0; y < texture2D.Description.Height; y++)
                {
                    // Copy a single line 
                    SharpDX.Utilities.CopyMemory(destPtr, sourcePtr, texture2D.Description.Width * 4);

                    // Advance pointers
                    sourcePtr = IntPtr.Add(sourcePtr, mapSource.RowPitch);
                    destPtr = IntPtr.Add(destPtr, mapDest.Stride);
                }

                // Release source and dest locks
                currentFrame.UnlockBits(mapDest);
                device.ImmediateContext.UnmapSubresource(texture2D, 0);

                screenResource.Dispose();
                duplicatedOutput.ReleaseFrame();

                return (GetDirectXFrameResult.Success, currentFrame);
            }
            catch (SharpDXException e)
            {
                if (e.ResultCode.Code == SharpDX.DXGI.ResultCode.WaitTimeout.Code)
                {
                    return (GetDirectXFrameResult.Timeout, null);
                }
                Logger.Write(e, "SharpDXException error.");
            }
            catch (Exception ex)
            {
                Logger.Write(ex);
            }
            return (GetDirectXFrameResult.Failure, null);
        }

        private void InitBitBlt()
        {
            try
            {
                _bitBltScreens.Clear();
                for (var i = 0; i < Screen.AllScreens.Length; i++)
                {
                    _bitBltScreens.Add(Screen.AllScreens[i].DeviceName, i);
                }
            }
            catch (Exception ex)
            {
                Logger.Write(ex);
            }
        }

        private void InitDirectX()
        {
            try
            {
                ClearDirectXOutputs();

                using var factory = new Factory1();
                foreach (var adapter in factory.Adapters1.Where(x => (x.Outputs?.Length ?? 0) > 0))
                {
                    foreach (var output in adapter.Outputs)
                    {
                        try
                        {
                            var device = new SharpDX.Direct3D11.Device(adapter);
                            var output1 = output.QueryInterface<Output1>();

                            var bounds = output1.Description.DesktopBounds;
                            var width = bounds.Right - bounds.Left;
                            var height = bounds.Bottom - bounds.Top;

                            // Create Staging texture CPU-accessible
                            var textureDesc = new Texture2DDescription
                            {
                                CpuAccessFlags = CpuAccessFlags.Read,
                                BindFlags = BindFlags.None,
                                Format = Format.B8G8R8A8_UNorm,
                                Width = width,
                                Height = height,
                                OptionFlags = ResourceOptionFlags.None,
                                MipLevels = 1,
                                ArraySize = 1,
                                SampleDescription = { Count = 1, Quality = 0 },
                                Usage = ResourceUsage.Staging
                            };

                            var texture2D = new Texture2D(device, textureDesc);

                            _directxScreens.Add(
                                output1.Description.DeviceName,
                                new DirectXOutput(adapter,
                                    device,
                                    output1.DuplicateOutput(device),
                                    texture2D,
                                    output1.Description.Rotation));
                        }
                        catch (Exception ex)
                        {
                            Logger.Write(ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Write(ex);
            }
        }

        private void RefreshCurrentScreenBounds()
        {
            CurrentScreenBounds = Screen.AllScreens[_bitBltScreens[SelectedScreen]].Bounds;
            CaptureFullscreen = true;
            NeedsInit = true;
            ScreenChanged?.Invoke(this, CurrentScreenBounds);
        }

        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            RefreshCurrentScreenBounds();
        }

        public static Bitmap ScaleImage(Bitmap pBmp, int pWidth, int pHeight)
        {
            try
            {
                Bitmap tmpBmp = new Bitmap(pWidth, pHeight);
                Graphics tmpG = Graphics.FromImage(tmpBmp);

                //tmpG.InterpolationMode = InterpolationMode.HighQualityBicubic;

                tmpG.DrawImage(pBmp,
                                           new Rectangle(0, 0, pWidth, pHeight),
                                           new Rectangle(0, 0, pBmp.Width, pBmp.Height),
                                           GraphicsUnit.Pixel);
                tmpG.Dispose();
                return tmpBmp;
            }
            catch
            {
                return null;
            }
        }

        private Bitmap ZoomImage(Bitmap bitmap, int destHeight, int destWidth)
        {
            try
            {
                System.Drawing.Image sourImage = bitmap;
                int width = 0, height = 0;
                //按比例縮放             
                int sourWidth = sourImage.Width;
                int sourHeight = sourImage.Height;
                if (sourHeight > destHeight || sourWidth > destWidth)
                {
                    if ((sourWidth * destHeight) > (sourHeight * destWidth))
                    {
                        width = destWidth;
                        height = (destWidth * sourHeight) / sourWidth;
                    }
                    else
                    {
                        height = destHeight;
                        width = (sourWidth * destHeight) / sourHeight;
                    }
                }
                else
                {
                    width = sourWidth;
                    height = sourHeight;
                }
                Bitmap destBitmap = new Bitmap(destWidth, destHeight);
                Graphics g = Graphics.FromImage(destBitmap);
                g.Clear(Color.Transparent);
                //設置畫布的描繪質量           
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(sourImage, new Rectangle((destWidth - width) / 2, (destHeight - height) / 2, width, height), 0, 0, sourImage.Width, sourImage.Height, GraphicsUnit.Pixel);
                g.Dispose();
                //設置壓縮質量       
                System.Drawing.Imaging.EncoderParameters encoderParams = new System.Drawing.Imaging.EncoderParameters();
                long[] quality = new long[1];
                quality[0] = 100;
                System.Drawing.Imaging.EncoderParameter encoderParam = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                encoderParams.Param[0] = encoderParam;
                sourImage.Dispose();
                return destBitmap;
            }
            catch
            {
                return bitmap;
            }
        }

        public Bitmap getnew(Image bit, double beishu)//beishu參數爲放大的倍數。放大縮小都可以，0.8即爲縮小至原來的0.8倍
        {
            Bitmap destBitmap = new Bitmap(Convert.ToInt32(bit.Width * beishu), Convert.ToInt32(bit.Height * beishu));
            Graphics g = Graphics.FromImage(destBitmap);
            g.Clear(Color.Transparent);
            //設置畫布的描繪質量           
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.DrawImage(bit, new Rectangle(0, 0, destBitmap.Width, destBitmap.Height), 0, 0, bit.Width, bit.Height, GraphicsUnit.Pixel);
            g.Dispose();
            return destBitmap;
        }
    }
}
