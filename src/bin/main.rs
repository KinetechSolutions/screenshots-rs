use std::error::Error;
use std::mem::{self, size_of};
use std::sync::mpsc::channel;
use std::time::Instant;
use std::{future, ptr};
use windows::core::{IInspectable, Interface, HSTRING, PCWSTR};
use windows::Foundation::TypedEventHandler;
use windows::Graphics::Capture::{Direct3D11CaptureFramePool, GraphicsCaptureItem};
use windows::Win32::Graphics::Direct3D::{
  D3D_DRIVER_TYPE, D3D_DRIVER_TYPE_HARDWARE, D3D_DRIVER_TYPE_WARP,
};
use windows::Win32::Graphics::Dxgi::Common::{
  DXGI_ALPHA_MODE_UNSPECIFIED, DXGI_FORMAT_B8G8R8A8_UNORM,
};
use windows::Win32::Graphics::Dxgi::{
  CreateDXGIFactory1, IDXGIDevice, IDXGIFactory1, IDXGIFactory2, IDXGIOutput1, IDXGISurface,
  DXGI_ERROR_UNSUPPORTED,
};
use windows::Win32::Graphics::Imaging::D2D::IWICImagingFactory2;
use windows::Win32::Graphics::Imaging::{
  CLSID_WICImagingFactory, GUID_WICPixelFormat32bppBGR, IWICImagingFactory,
  WICBitmapDitherTypeNone, WICBitmapPaletteTypeMedianCut,
};

use windows::Graphics::DirectX::DirectXPixelFormat;
use windows::Graphics::Imaging::{
  BitmapAlphaMode, BitmapEncoder, BitmapPixelFormat, BitmapProperties,
};
use windows::Storage::{CreationCollisionOption, FileAccessMode, StorageFolder};
use windows::Win32::Foundation::{
  BOOL, D2DERR_BITMAP_BOUND_AS_TARGET, D2DERR_BITMAP_CANNOT_DRAW, ERROR_INVALID_DATA, HWND, LPARAM,
};

use windows::Win32::Graphics::Direct3D11::{
  D3D11CreateDevice, ID3D11Device, ID3D11Resource, ID3D11Texture2D, D3D11_BIND_FLAG,
  D3D11_CPU_ACCESS_READ, D3D11_CREATE_DEVICE_BGRA_SUPPORT, D3D11_CREATE_DEVICE_DEBUG,
  D3D11_MAP_READ, D3D11_RESOURCE_MISC_FLAG, D3D11_SDK_VERSION, D3D11_TEXTURE2D_DESC,
  D3D11_USAGE_STAGING,
};

use windows::Win32::System::WinRT::Graphics::Capture::IGraphicsCaptureItemInterop;
use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED, RO_INIT_TYPE};

use windows::Win32::Graphics::Dwm::{DwmGetWindowAttribute, DWMWA_CLOAKED, DWM_CLOAKED_SHELL};

use windows::Win32::Graphics::Gdi::{
  BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, CreateDCW, CreatedHDC, DeleteDC,
  GetDeviceCaps, GetMonitorInfoW, GetWindowDC, MonitorFromWindow, SelectObject, HBITMAP, HDC,
  HMONITOR, HORZRES, MONITORINFOEXW, MONITOR_DEFAULTTONEAREST, VERTRES,
};

use screenshots::d3d;
use windows::Win32::UI::WindowsAndMessaging::{
  EnumWindows, GetAncestor, GetClassNameW, GetShellWindow, GetWindowInfo, GetWindowLongW,
  GetWindowTextW, IsWindowVisible, GA_ROOT, GWL_EXSTYLE, GWL_STYLE, SWP_NOACTIVATE, SWP_NOZORDER,
  WINDOWINFO, WS_DISABLED, WS_EX_TOOLWINDOW,
};

fn main() -> windows::core::Result<()> {
  // futures::executor::block_on(entrypoint()) // TEMP DISABLE
  entrypoint()
}

fn entrypoint() -> windows::core::Result<()> {
  let start = Instant::now();
  unsafe {
    EnumWindows(Some(enum_window), LPARAM(0));
    println!("finished screenshot routine in {:?}", start.elapsed());
    Ok(())
  }
}

struct Capture {
  window_ctx: HDC,
  name: String,
  bitmap: HBITMAP,
  w: i32,
  h: i32,
  gdi_buffer: Vec<u8>,
}

impl Capture {
  pub unsafe fn capture(
    name: String,
    handle: HWND,
    w: i32,
    h: i32,
  ) -> Result<Capture, Box<dyn Error>> {
    let ctx: HDC = GetWindowDC(handle);
    let mem_context: CreatedHDC = CreateCompatibleDC(ctx);
    if mem_context.is_invalid() {
      return Err("failed to capture graphics display!".into());
    }

    println!("{},{}", w, h);
    let bitmap: HBITMAP = CreateCompatibleBitmap(ctx, w, h);
    if bitmap.is_invalid() {
      DeleteDC(mem_context);
      DeleteDC(mem_context);
      return Err("Can't create a Windows buffer".into());
    }

    let res = SelectObject(mem_context, bitmap);
    if res.is_invalid() {
      DeleteDC(mem_context);
      DeleteDC(mem_context);
      return Err(Box::new(std::io::Error::from_raw_os_error(
        ERROR_INVALID_DATA.0 as _,
      )));
    }

    Ok(Capture {
      window_ctx: ctx,
      name,
      bitmap,
      w,
      h,
      gdi_buffer: Vec::new(),
    })
  }

  pub fn frame(&self, buffer: &mut Vec<u8>) -> Result<(), Box<dyn std::error::Error>> {
    unsafe {
      let res = BitBlt(
        self.window_ctx,
        0,
        0,
        self.w,
        self.h,
        self.window_ctx,
        0,
        0,
        windows::Win32::Graphics::Gdi::ROP_CODE(SWP_NOACTIVATE.0 | SWP_NOZORDER.0),
      );
      if !res.as_bool() {
        return Err("Failed to copy screen to Windows buffer".into());
      }

      if self.w == 0 || self.h == 0 {
        return Ok(());
      }

      let stride = self.w * 4;
      let size: usize = (stride * self.h) as usize;
      let mut data1: Vec<u8> = Vec::with_capacity(size);
      data1.set_len(size);
      buffer.resize(size, 0);

      Ok(())
    }
  }
}

extern "system" fn enum_window(window: HWND, _state: LPARAM) -> BOOL {
  let mut inf: WINDOWINFO = WINDOWINFO {
    cbSize: core::mem::size_of::<WINDOWINFO>() as u32,
    ..Default::default()
  };
  let info = WindowInfo::new(window);
  unsafe { GetWindowInfo(window, &mut inf) };
  let monitor = unsafe { MonitorFromWindow(window, MONITOR_DEFAULTTONEAREST) };

  println!("capture {}, {}", info.title_str, monitor.0);

  true.into()
}

async fn take_screenshot(device: GraphicsCaptureItem) -> windows::core::Result<()> {
  let mut ndisplay = 0;
  let mut factory = ptr::null_mut();
  let dxfactory: IDXGIFactory2 = unsafe { CreateDXGIFactory1()? };
  let dxdevice = device.cast::<IDXGIDevice>()?;

  let factory = factory as *mut IDXGIFactory1;
  let mut adapter = ptr::null_mut();
  unsafe {
    (*factory).EnumAdapters1(*adapter);
  }

  // let frame_pool = Direct3D11CaptureFramePool::CreateFreeThreaded(
  //   device,
  //   DirectXPixelFormat::B8G8R8A8Typeless,
  //   1,
  //   &device.Size().unwrap(),
  // )?;
  // let session = frame_pool.CreateCaptureSession(device)?;

  // let (sender, receiver) = channel();
  // frame_pool.FrameArrived(
  //   &TypedEventHandler::<Direct3D11CaptureFramePool, IInspectable>::new({
  //     move |frame_pool, _| {
  //       let frame_pool = frame_pool.as_ref().unwrap();
  //       let frame = frame_pool.TryGetNextFrame()?;
  //       sender.send(frame).unwrap();
  //       Ok(())
  //     }
  //   }),
  // )?;
  // session.StartCapture()?;

  let device = create_device().unwrap();
  let mut desc = std::ptr::null_mut();
  let surface: IDXGISurface = unsafe {
    (*factory)
      .CreateSwapChain(dxdevice, desc)
      .unwrap()
      .GetBuffer(0)?
  };
  if adapter.is_null() {
    return Ok(());
  }

  // Otherwise, we get the next output of the current adapter.

  let output = unsafe {
    let mut output = ptr::null_mut();
    (*factory).EnumAdapters(*adapter)?.EnumOutputs(*output);
    output
  };

  // If the current adapter is done, we free it.
  // We return None so the caller gets the next adapter and tries again.

  if output.is_null() {
    adapter = ptr::null_mut();
    return Ok(());
  }

  let desc = unsafe { surface.GetDesc() };

  // We cast it up to the version needed for desktop duplication.
  let mut inner = unsafe { (*output) as *mut IDXGIOutput1 };
  // let props = BitmapProperties::from(D2DERR_BITMAP_BOUND_AS_TARGET | D2DERR_BITMAP_CANNOT_DRAW) {
  //   pixelFormat: D2D1_PIXEL_FORMAT {
  //     format: DXGI_FORMAT_B8G8R8A8_UNORM,
  //     alphaMode: DXGI_ALPHA_MODE_UNSPECIFIED,
  //   },
  //   dpiX: 96.0,
  //   dpiY: 96.0,
  //   bitmapOptions: BitmapO,
  //   ..Default::default()
  // };

  // unsafe {
  //   let bitmap = target.CreateBitmapFromDxgiSurface(&surface, Some(&props))?;
  //   target.SetTarget(&bitmap);
  // };
  // let texture = unsafe {
  //   // let frame = receiver.recv().unwrap();
  //   let source_texture: ID3D11Texture2D = d3d::get_d3d_interface_from_object(&device.Surface()?)?;
  //   let mut desc = D3D11_TEXTURE2D_DESC::default();
  //   source_texture.GetDesc(&mut desc);
  //   desc.BindFlags = D3D11_BIND_FLAG(0);
  //   desc.MiscFlags = D3D11_RESOURCE_MISC_FLAG(0);
  //   desc.Usage = D3D11_USAGE_STAGING;
  //   desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
  //   let copy_texture = { device.CreateTexture2D(&desc, std::ptr::null())? };

  //   // d3d_context.CopyResource(Some(copy_texture.cast()?), Some(source_texture.cast()?));

  //   // session.Close()?;
  //   // frame_pool.Close()?;

  //   copy_texture
  // };

  // let bits = unsafe {
  //   let mut desc = D3D11_TEXTURE2D_DESC::default();
  //   texture.GetDesc(&mut desc);

  //   let resource: ID3D11Resource = texture.cast()?;
  //   // let mapped = d3d_context.Map(&resource, 0, D3D11_MAP_READ, 0)?;

  //   let slice: &[u8] = {
  //     std::slice::from_raw_parts(
  //       mapped.pData as *const _,
  //       (desc.Height * mapped.RowPitch) as usize,
  //     )
  //   };

  //   let bytes_per_pixel = ARGBPIXEL_BYTE_LEN as u32;
  //   let mut bits = vec![0u8; (desc.Width * desc.Height * bytes_per_pixel) as usize];
  //   for row in 0..desc.Height {
  //     let data_begin = (row * (desc.Width * bytes_per_pixel)) as usize;
  //     let data_end = ((row + 1) * (desc.Width * bytes_per_pixel)) as usize;
  //     let slice_begin = (row * mapped.RowPitch) as usize;
  //     let slice_end = slice_begin + (desc.Width * bytes_per_pixel) as usize;
  //     bits[data_begin..data_end].copy_from_slice(&slice[slice_begin..slice_end]);
  //   }

  //   // d3d_context.Unmap(Some(resource), 0);

  //   bits
  // };

  // let path = std::env::current_dir()
  //   .unwrap()
  //   .to_string_lossy()
  //   .to_string();
  // let folder = StorageFolder::GetFolderFromPathAsync(HSTRING::from(path))?.await?;
  // let file = folder
  //   .CreateFileAsync(
  //     &HSTRING::from(format!("{:?}.png", item.DisplayName())),
  //     CreationCollisionOption::ReplaceExisting,
  //   )?
  //   .await?;

  // {
  //   let stream = file.OpenAsync(FileAccessMode::ReadWrite)?.await?;
  //   let encoder = BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId()?, &stream)?.await?;
  //   encoder.SetPixelData(
  //     BitmapPixelFormat::Bgra8,
  //     BitmapAlphaMode::Premultiplied,
  //     item_size.Width as u32,
  //     item_size.Height as u32,
  //     1.0,
  //     1.0,
  //     &bits,
  //   )?;

  //   encoder.FlushAsync()?.await?;
  // }

  Ok(())
}

fn create_device() -> Option<ID3D11Device> {
  let device = std::ptr::null_mut();
  let mut result = create_device_with_type(D3D_DRIVER_TYPE_HARDWARE, device);

  if let Some(device) = result {
    // if err. == DXGI_ERROR_UNSUPPORTED {
    //   result = create_device_with_type(D3D_DRIVER_TYPE_WARP);
    // }
    return Some(device);
  }

  None
}

fn create_device_with_type(
  drive_type: D3D_DRIVER_TYPE,
  device_ptr: *mut Option<ID3D11Device>,
) -> Option<ID3D11Device> {
  let mut flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;

  if cfg!(debug_assertions) {
    flags |= D3D11_CREATE_DEVICE_DEBUG;
  }

  let device = std::ptr::null_mut();
  unsafe {
    d3d::create_d3d_device_with_type(drive_type, flags, device).unwrap();
    device_ptr.replace((*device).take())
  }
}

#[derive(Clone)]
pub struct WindowInfo {
  pub handle: HWND,
  pub title_str: String,
  pub class_name: String,
  pub title_raw: Vec<u16>,
}

impl WindowInfo {
  pub fn new(window_handle: HWND) -> Self {
    unsafe {
      let mut title = [0u16; 512];
      GetWindowTextW(window_handle, &mut title);
      let mut title_str = String::from_utf16_lossy(&title);
      truncate_to_first_null_char(&mut title_str);

      let mut class_name = [0u16; 512];
      GetClassNameW(window_handle, &mut class_name);
      let mut class_name = String::from_utf16_lossy(&class_name);
      truncate_to_first_null_char(&mut class_name);

      Self {
        handle: window_handle,
        title_str,
        class_name,
        title_raw: title.to_vec(),
      }
    }
  }

  pub fn matches(&self, title: &str, class_name: &str) -> bool {
    self.title_str == title && self.class_name == class_name
  }
}

fn truncate_to_first_null_char(input: &mut String) {
  if let Some(index) = input.find('\0') {
    input.truncate(index);
  }
}

#[derive(Clone)]
pub struct DisplayInfo {
  pub handle: HMONITOR,
  pub display_name: String,
}

impl DisplayInfo {
  pub fn new(monitor_handle: HMONITOR) -> windows::core::Result<Self> {
    let mut info = MONITORINFOEXW::default();
    info.monitorInfo.cbSize = std::mem::size_of::<MONITORINFOEXW>() as u32;

    unsafe {
      GetMonitorInfoW(monitor_handle, &mut info as *mut _ as *mut _).ok()?;
    }

    let display_name = String::from_utf16_lossy(&info.szDevice)
      .trim_matches(char::from(0))
      .to_string();

    Ok(Self {
      handle: monitor_handle,
      display_name,
    })
  }
}

pub trait CaptureWindowCandidate {
  fn is_capturable_window(&self) -> bool;
}

impl CaptureWindowCandidate for WindowInfo {
  fn is_capturable_window(&self) -> bool {
    unsafe {
      if self.title_str.is_empty()
        || self.handle == GetShellWindow()
        || IsWindowVisible(self.handle).as_bool() == false
        || GetAncestor(self.handle, GA_ROOT) != self.handle
      {
        return false;
      }

      let style = GetWindowLongW(self.handle, GWL_STYLE);
      if style & (WS_DISABLED.0 as i32) == 1 {
        return false;
      }

      // No tooltips
      let ex_style = GetWindowLongW(self.handle, GWL_EXSTYLE);
      if ex_style & (WS_EX_TOOLWINDOW.0 as i32) == 1 {
        return false;
      }

      // Check to see if the self is cloaked if it's a UWP
      if self.class_name == "Windows.UI.Core.CoreWindow"
        || self.class_name == "ApplicationFrameWindow"
      {
        let mut cloaked: u32 = 0;
        if DwmGetWindowAttribute(
          self.handle,
          DWMWA_CLOAKED,
          &mut cloaked as *mut _ as *mut _,
          std::mem::size_of::<u32>() as u32,
        )
        .is_ok()
          && cloaked == DWM_CLOAKED_SHELL
        {
          return false;
        }
      }

      // Unfortunate work-around. Not sure how to avoid this.
      if is_known_blocked_window(&self) {
        return false;
      }
    }
    true
  }
}

fn is_known_blocked_window(window_info: &WindowInfo) -> bool {
  // Task View
  window_info.matches("Task View", "Windows.UI.Core.CoreWindow") ||
  // XAML Islands
  window_info.matches("DesktopWindowXamlSource", "Windows.UI.Core.CoreWindow") ||
  // XAML Popups
  window_info.matches("PopupHost", "Xaml_WindowedPopupClass")
}
