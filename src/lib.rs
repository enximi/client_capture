//! # 客户区域截图
//! 更简单的捕获一个窗口客户区域的图像。
//! 通过指定窗口类名和窗口标题，可以捕获到指定窗口的客户区域图像。
//! 如果窗口不存在或者窗口被关闭，会自动重试。
//! 是 [windows_capture](https://github.com/NiiightmareXD/windows-capture) 的封装。

use std::time::{Duration, Instant};

use anyhow::{anyhow, Result};
use image::{DynamicImage, RgbaImage};
use tokio::spawn;
use tokio::sync::watch;
use tokio::task::JoinHandle;
use tokio::time::sleep;
use window_inspector::find::get_hwnd_ref_cache;
use window_inspector::position_size::{get_client_xywh, get_window_xywh_exclude_shadow};
use windows_capture::window::Window;
use windows_capture::{
    capture::GraphicsCaptureApiHandler,
    frame::Frame,
    graphics_capture_api::InternalCaptureControl,
    settings::{ColorFormat, CursorCaptureSettings, DrawBorderSettings, Settings},
};

struct CaptureMessage {
    // 暂停watch
    pause_rx: watch::Receiver<bool>,
    // 停止watch
    stop_rx: watch::Receiver<bool>,
    // 窗口句柄
    hwnd: isize,
    // 额外需要去除的边框：左上右下
    border: (u32, u32, u32, u32),
    // 图像发送者
    img_tx: watch::Sender<Option<(DynamicImage, Instant)>>,
}

struct Capture {
    message: CaptureMessage,
}

impl Capture {
    fn to_img(&self, frame: &mut Frame) -> Result<DynamicImage> {
        let window_xywh = get_window_xywh_exclude_shadow(self.message.hwnd)?;
        let frame_wh = (frame.width(), frame.height());
        if window_xywh.2 != frame_wh.0 || window_xywh.3 != frame_wh.1 {
            return Err(anyhow!("窗口大小与帧大小不一致"));
        }
        let client_xywh = get_client_xywh(self.message.hwnd)?;
        if client_xywh.2 == 0 || client_xywh.3 == 0 {
            return Err(anyhow!("窗口大小为0"));
        }
        if client_xywh.0 < window_xywh.0
            || client_xywh.1 < window_xywh.1
            || client_xywh.0 + client_xywh.2 as i32 > window_xywh.0 + window_xywh.2 as i32
            || client_xywh.1 + client_xywh.3 as i32 > window_xywh.1 + window_xywh.3 as i32
        {
            return Err(anyhow!("客户区域超出窗口范围"));
        };
        let client_xy_in_window = (
            (client_xywh.0 - window_xywh.0) as u32,
            (client_xywh.1 - window_xywh.1) as u32,
        );
        if client_xywh.2 < self.message.border.0 + self.message.border.2
            || client_xywh.3 < self.message.border.1 + self.message.border.3
        {
            return Err(anyhow!("客户区域小于边框大小"));
        }
        let buffer = frame
            .buffer_crop(
                client_xy_in_window.0 + self.message.border.0,
                client_xy_in_window.1 + self.message.border.1,
                client_xy_in_window.0 + client_xywh.2 - self.message.border.2,
                client_xy_in_window.1 + client_xywh.3 - self.message.border.3,
            )?
            .as_raw_nopadding_buffer()?
            .to_vec();
        let img = RgbaImage::from_raw(client_xywh.2, client_xywh.3, buffer)
            .ok_or_else(|| anyhow!("转换为RgbaImage失败"))?;
        Ok(DynamicImage::ImageRgba8(img))
    }
}

impl GraphicsCaptureApiHandler for Capture {
    type Flags = CaptureMessage;

    type Error = Box<dyn std::error::Error + Send + Sync>;

    fn new(message: Self::Flags) -> Result<Self, Self::Error> {
        Ok(Self { message })
    }

    fn on_frame_arrived(
        &mut self,
        frame: &mut Frame,
        capture_control: InternalCaptureControl,
    ) -> Result<(), Self::Error> {
        // 停止
        if *self.message.stop_rx.borrow() {
            capture_control.stop();
            return Ok(());
        }

        // 暂停
        if *self.message.pause_rx.borrow() {
            return Ok(());
        }

        match self.to_img(frame) {
            Ok(img) => {
                if let Err(e) = self.message.img_tx.send(Some((img, Instant::now()))) {
                    log::warn!("发送图像失败: {}", e);
                    return Ok(());
                }
            }
            Err(e) => {
                log::warn!("转换图像失败: {}", e);
                return Ok(());
            }
        };

        Ok(())
    }

    fn on_closed(&mut self) -> Result<(), Self::Error> {
        log::debug!("捕获对象已不存在");
        Ok(())
    }
}

pub struct ClientCapture {
    window_class: String,
    window_title: String,
    border: (u32, u32, u32, u32),
    pause_tx: watch::Sender<bool>,
    pause_rx: watch::Receiver<bool>,
    stop_tx: watch::Sender<bool>,
    stop_rx: watch::Receiver<bool>,
    img_tx: watch::Sender<Option<(DynamicImage, Instant)>>,
    img_rx: watch::Receiver<Option<(DynamicImage, Instant)>>,
    capture_handle: Option<JoinHandle<()>>,
    /// 可以接受的截图延时
    delay: Duration,
}

impl ClientCapture {
    /// 创建一个新的截图对象
    /// # 参数
    /// - window_class: 窗口类名
    /// - window_title: 窗口标题
    /// - border: 在客户区域的基础上额外需要去除的边框：左上右下
    /// - delay: 可以接受的截图延时，默认50ms
    pub fn new(
        window_class: String,
        window_title: String,
        border: Option<(u32, u32, u32, u32)>,
        delay: Option<Duration>,
    ) -> Self {
        let (pause_tx, pause_rx) = watch::channel(false);
        let (stop_tx, stop_rx) = watch::channel(false);
        let (img_tx, img_rx) = watch::channel(None);
        Self {
            window_class,
            window_title,
            border: border.unwrap_or((0, 0, 0, 0)),
            pause_tx,
            pause_rx,
            stop_tx,
            stop_rx,
            img_tx,
            img_rx,
            capture_handle: None,
            delay: delay.unwrap_or(Duration::from_millis(50)),
        }
    }

    /// 启动截图线程
    /// # 返回
    /// - Ok: 成功
    /// - Err: 如果截图线程正在运行，则返回错误
    pub fn start(&mut self) -> Result<()> {
        if self.is_running() {
            return Err(anyhow!("截图线程正在运行"));
        }
        self.stop_tx.send(false).unwrap();
        self.pause_tx.send(false).unwrap();
        let window_class = self.window_class.clone();
        let window_title = self.window_title.clone();
        let pause_rx = self.pause_rx.clone();
        let stop_rx = self.stop_rx.clone();
        let img_tx = self.img_tx.clone();
        let border = self.border;
        let capture_handle = spawn(async move {
            loop {
                if *stop_rx.borrow() {
                    break;
                }
                match get_hwnd_ref_cache(&window_class, &window_title) {
                    Ok(hwnd) => {
                        let message = CaptureMessage {
                            pause_rx: pause_rx.clone(),
                            stop_rx: stop_rx.clone(),
                            hwnd,
                            border,
                            img_tx: img_tx.clone(),
                        };
                        let window = Window::from_raw_hwnd(hwnd);
                        let settings = Settings::new(
                            window,
                            CursorCaptureSettings::Default,
                            DrawBorderSettings::Default,
                            ColorFormat::Rgba8,
                            message,
                        );
                        if let Err(e) = Capture::start(settings) {
                            log::warn!("截图失败: {}", e);
                            sleep(Duration::from_millis(500)).await;
                        }
                        log::info!("截图线程结束");
                    }
                    Err(e) => {
                        log::warn!("获取窗口句柄失败: {}", e);
                        sleep(Duration::from_millis(500)).await;
                    }
                }
            }
        });
        self.capture_handle = Some(capture_handle);
        Ok(())
    }

    /// 是否正在运行
    pub fn is_running(&self) -> bool {
        if let Some(handle) = self.capture_handle.as_ref() {
            !handle.is_finished()
        } else {
            false
        }
    }

    /// 暂停截图，在不需要截图的时候可以暂停，也许可以减少资源占用
    pub fn pause(&self) {
        self.pause_tx.send(true).unwrap();
    }

    /// 恢复截图
    pub fn resume(&self) {
        self.pause_tx.send(false).unwrap();
    }

    /// 停止截图，把Capture线程关掉
    pub fn stop(&self) {
        self.stop_tx.send(true).unwrap();
    }

    /// 获取图像
    /// # 返回
    /// - Ok: 图像
    /// - Err: 图像为空或者截图时间距离现在超过50ms
    pub fn get_img(&mut self) -> Result<DynamicImage> {
        let img = self.img_rx.borrow().clone();
        if let Some((img, time)) = img {
            if time.elapsed() > self.delay {
                Err(anyhow!("图像已过期"))
            } else {
                Ok(img)
            }
        } else {
            Err(anyhow!("图像为空"))
        }
    }
}

impl Drop for ClientCapture {
    fn drop(&mut self) {
        self.stop();
        if let Some(handle) = self.capture_handle.take() {
            spawn(async move {
                if let Err(e) = handle.await {
                    log::warn!("截图线程异常结束: {}", e);
                }
            });
        }
    }
}
