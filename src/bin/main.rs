use std::sync::mpsc::channel;
use std::time::Instant;
use windows::core::{IInspectable, Interface, HSTRING};
use windows::Foundation::TypedEventHandler;
use windows::Graphics::Capture::{Direct3D11CaptureFramePool, GraphicsCaptureItem};

use windows::Graphics::DirectX::DirectXPixelFormat;
use windows::Graphics::Imaging::{BitmapAlphaMode, BitmapEncoder, BitmapPixelFormat};
use windows::Storage::{CreationCollisionOption, FileAccessMode, StorageFolder};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM};

use windows::Win32::Graphics::Direct3D11::{
  ID3D11Resource, ID3D11Texture2D, D3D11_BIND_FLAG, D3D11_CPU_ACCESS_READ, D3D11_MAP_READ,
  D3D11_RESOURCE_MISC_FLAG, D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING,
};

use windows::Win32::System::WinRT::Graphics::Capture::IGraphicsCaptureItemInterop;
use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED};

use windows::Win32::Graphics::Dwm::{DwmGetWindowAttribute, DWMWA_CLOAKED, DWM_CLOAKED_SHELL};

use windows::Win32::Graphics::Gdi::{GetMonitorInfoW, HMONITOR, MONITORINFOEXW};

use screenshots::d3d;
use windows::Win32::UI::WindowsAndMessaging::{
  EnumWindows, GetAncestor, GetClassNameW, GetShellWindow, GetWindowInfo, GetWindowLongW,
  GetWindowTextW, IsWindowVisible, GA_ROOT, GWL_EXSTYLE, GWL_STYLE, WINDOWINFO, WS_DISABLED,
  WS_EX_TOOLWINDOW,
};

fn create_capture_item_for_window(
  window_handle: HWND,
) -> windows::core::Result<GraphicsCaptureItem> {
  let interop = windows::core::factory::<GraphicsCaptureItem, IGraphicsCaptureItemInterop>()?;
  unsafe { interop.CreateForWindow(window_handle) }
}

pub trait CaptureWindowCandidate {
  fn is_capturable_window(&self) -> bool;
}

impl CaptureWindowCandidate for WindowInfo {
  fn is_capturable_window(&self) -> bool {
    unsafe {
      if self.title.is_empty()
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
    }
    true
  }
}

fn main() -> windows::core::Result<()> {
  let start = Instant::now();
  unsafe {
    RoInitialize(RO_INIT_MULTITHREADED)?;
    EnumWindows(Some(enum_window), LPARAM(0));
    println!("finished screenshot routine in {:?}", start.elapsed());
    Ok(())
  }
}

extern "system" fn enum_window(window: HWND, _state: LPARAM) -> BOOL {
  unsafe {
    let mut info = WINDOWINFO {
      cbSize: core::mem::size_of::<WINDOWINFO>() as u32,
      ..Default::default()
    };
    GetWindowInfo(window, &mut info).unwrap();
    let gpu_item = create_capture_item_for_window(window).unwrap();
    take_screenshot(&gpu_item).expect("failed to snap screenshot from window.");

    true.into()
  }
}

fn take_screenshot(item: &GraphicsCaptureItem) -> windows::core::Result<()> {
  let item_size = item.Size()?;

  let d3d_device = d3d::create_d3d_device()?;
  let d3d_context = unsafe {
    let mut d3d_context = None;
    d3d_device.GetImmediateContext(&mut d3d_context);
    d3d_context.unwrap()
  };
  let device = d3d::create_direct3d_device(&d3d_device)?;
  let frame_pool = Direct3D11CaptureFramePool::CreateFreeThreaded(
    &device,
    DirectXPixelFormat::R32UInt,
    1,
    &item_size,
  )?;
  let session = frame_pool.CreateCaptureSession(item)?;

  let (sender, receiver) = channel();
  frame_pool.FrameArrived(
    &TypedEventHandler::<Direct3D11CaptureFramePool, IInspectable>::new({
      move |frame_pool, _| {
        let frame_pool = frame_pool.as_ref().unwrap();
        let frame = frame_pool.TryGetNextFrame()?;
        sender.send(frame).unwrap();
        Ok(())
      }
    }),
  )?;
  session.StartCapture()?;

  let texture = unsafe {
    let frame = receiver.recv().unwrap();

    let source_texture: ID3D11Texture2D = d3d::get_d3d_interface_from_object(&frame.Surface()?)?;
    let mut desc = D3D11_TEXTURE2D_DESC::default();
    source_texture.GetDesc(&mut desc);
    desc.BindFlags = D3D11_BIND_FLAG(0);
    desc.MiscFlags = D3D11_RESOURCE_MISC_FLAG(0);
    desc.Usage = D3D11_USAGE_STAGING;
    desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
    let copy_texture = { d3d_device.CreateTexture2D(&desc, std::ptr::null())? };

    d3d_context.CopyResource(Some(copy_texture.cast()?), Some(source_texture.cast()?));

    session.Close()?;
    frame_pool.Close()?;

    copy_texture
  };

  let bits = unsafe {
    let mut desc = D3D11_TEXTURE2D_DESC::default();
    texture.GetDesc(&mut desc);

    let resource: ID3D11Resource = texture.cast()?;
    let mapped = d3d_context.Map(&resource, 0, D3D11_MAP_READ, 0)?;

    let slice: &[u8] = {
      std::slice::from_raw_parts(
        mapped.pData as *const _,
        (desc.Height * mapped.RowPitch) as usize,
      )
    };

    let bytes_per_pixel = 4;
    let mut bits = vec![0u8; (desc.Width * desc.Height * bytes_per_pixel) as usize];
    for row in 0..desc.Height {
      let data_begin = (row * (desc.Width * bytes_per_pixel)) as usize;
      let data_end = ((row + 1) * (desc.Width * bytes_per_pixel)) as usize;
      let slice_begin = (row * mapped.RowPitch) as usize;
      let slice_end = slice_begin + (desc.Width * bytes_per_pixel) as usize;
      bits[data_begin..data_end].copy_from_slice(&slice[slice_begin..slice_end]);
    }

    d3d_context.Unmap(Some(resource), 0);

    bits
  };

  let path = std::env::current_dir()
    .unwrap()
    .to_string_lossy()
    .to_string();
  let folder = StorageFolder::GetFolderFromPathAsync(HSTRING::from(path))?.get()?;
  let file = folder
    .CreateFileAsync(
      &HSTRING::from(format!("{:?}.png", item.DisplayName())),
      CreationCollisionOption::ReplaceExisting,
    )?
    .get()?;

  {
    let stream = file.OpenAsync(FileAccessMode::ReadWrite)?.get()?;
    let encoder = BitmapEncoder::CreateAsync(BitmapEncoder::PngEncoderId()?, &stream)?.get()?;
    encoder.SetPixelData(
      BitmapPixelFormat::Bgra8,
      BitmapAlphaMode::Premultiplied,
      item_size.Width as u32,
      item_size.Height as u32,
      1.0,
      1.0,
      &bits,
    )?;

    encoder.FlushAsync()?.get()?;
  }

  Ok(())
}

#[derive(Clone)]
pub struct WindowInfo {
  pub handle: HWND,
  pub title: String,
  pub class_name: String,
}

impl WindowInfo {
  pub fn new(window_handle: HWND) -> Self {
    unsafe {
      let mut title = [0u16; 512];
      GetWindowTextW(window_handle, &mut title);
      let mut title = String::from_utf16_lossy(&title);
      truncate_to_first_null_char(&mut title);

      let mut class_name = [0u16; 512];
      GetClassNameW(window_handle, &mut class_name);
      let mut class_name = String::from_utf16_lossy(&class_name);
      truncate_to_first_null_char(&mut class_name);

      Self {
        handle: window_handle,
        title,
        class_name,
      }
    }
  }

  pub fn matches(&self, title: &str, class_name: &str) -> bool {
    self.title == title && self.class_name == class_name
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
