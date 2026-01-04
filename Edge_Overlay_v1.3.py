import sys
import cv2
import numpy as np
import dxcam
from PyQt5 import QtCore, QtGui, QtWidgets
import win32gui
import win32con
import win32api
import ctypes
import keyboard

# High DPI 인식 설정 (High DPI Awareness)
try:
    # Process_Per_Monitor_DPI_Aware_V2 = 2
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

class Config:
    CANNY_MIN = 50
    CANNY_MAX = 150
    EDGE_COLOR = (0, 255, 0)  # 초록색 (RGB)
    EDGE_THICKNESS = 1
    EDGE_OPACITY = 255
    REFRESH_RATE = 60
    AUTO_SIGMA = 0.33
    REALTIME_AUTO = False
    
class SettingsWidget(QtWidgets.QWidget):
    valueChanged = QtCore.pyqtSignal()
    sig_req_auto_adjust = QtCore.pyqtSignal()
    sig_req_window_select = QtCore.pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Edge Overlay v1.3")
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self.resize(300, 400)
        
        main_layout = QtWidgets.QVBoxLayout()

        # 그룹 1: 자동 설정 (Auto Settings)
        group_auto = QtWidgets.QGroupBox("자동 설정 (Auto Settings)")
        layout_auto = QtWidgets.QVBoxLayout()

        # Sigma
        self.sigma_label = QtWidgets.QLabel(f"Sigma: {Config.AUTO_SIGMA:.2f}")
        layout_auto.addWidget(self.sigma_label)
        
        self.sigma_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.sigma_slider.setRange(0, 100) # 0.00 ~ 1.00
        self.sigma_slider.setValue(int(Config.AUTO_SIGMA * 100))
        self.sigma_slider.valueChanged.connect(self.update_config)
        layout_auto.addWidget(self.sigma_slider)

        # 자동 조절 버튼 (Run Once)
        self.auto_btn = QtWidgets.QPushButton("자동 조절 (Run Once)")
        self.auto_btn.clicked.connect(self.sig_req_auto_adjust.emit)
        layout_auto.addWidget(self.auto_btn)

        # 실시간 자동 조절 체크박스 (Real-time)
        self.realtime_chk = QtWidgets.QCheckBox("실시간 자동 조절 (Real-time)")
        self.realtime_chk.setChecked(Config.REALTIME_AUTO)
        self.realtime_chk.stateChanged.connect(self.update_config)
        layout_auto.addWidget(self.realtime_chk)

        group_auto.setLayout(layout_auto)
        main_layout.addWidget(group_auto)

        # 그룹 2: 임계값 수동 설정 (Manual Thresholds)
        group_thresh = QtWidgets.QGroupBox("임계값 수동 설정 (Manual Thresholds)")
        layout_thresh = QtWidgets.QVBoxLayout()
        
        # Canny Min
        self.canny_min_label = QtWidgets.QLabel(f"Canny Min: {Config.CANNY_MIN}")
        layout_thresh.addWidget(self.canny_min_label)
        self.canny_min_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.canny_min_slider.setRange(0, 255)
        self.canny_min_slider.setValue(Config.CANNY_MIN)
        self.canny_min_slider.valueChanged.connect(self.update_config)
        layout_thresh.addWidget(self.canny_min_slider)
        
        # Canny Max
        self.canny_max_label = QtWidgets.QLabel(f"Canny Max: {Config.CANNY_MAX}")
        layout_thresh.addWidget(self.canny_max_label)
        self.canny_max_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.canny_max_slider.setRange(0, 255)
        self.canny_max_slider.setValue(Config.CANNY_MAX)
        self.canny_max_slider.valueChanged.connect(self.update_config)
        layout_thresh.addWidget(self.canny_max_slider)

        group_thresh.setLayout(layout_thresh)
        main_layout.addWidget(group_thresh)

        # 그룹 3: 모양 설정 (Appearance)
        group_appear = QtWidgets.QGroupBox("모양 설정 (Appearance)")
        layout_appear = QtWidgets.QVBoxLayout()

        # 두께 (Thickness)
        layout_appear.addWidget(QtWidgets.QLabel("두께 (Thickness)"))
        self.thickness_spin = QtWidgets.QSpinBox()
        self.thickness_spin.setRange(1, 5)
        self.thickness_spin.setValue(Config.EDGE_THICKNESS)
        self.thickness_spin.valueChanged.connect(self.update_config)
        layout_appear.addWidget(self.thickness_spin)
        
        # 투명도 (Opacity)
        self.opacity_label = QtWidgets.QLabel(f"투명도 (Opacity): {Config.EDGE_OPACITY}")
        layout_appear.addWidget(self.opacity_label)
        self.opacity_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.opacity_slider.setRange(0, 255)
        self.opacity_slider.setValue(Config.EDGE_OPACITY)
        self.opacity_slider.valueChanged.connect(self.update_config)
        layout_appear.addWidget(self.opacity_slider)

        # 색상 변경 버튼 (Color)
        self.color_btn = QtWidgets.QPushButton("색상 변경 (Color)")
        self.color_btn.clicked.connect(self.change_color)
        layout_appear.addWidget(self.color_btn)

        group_appear.setLayout(layout_appear)
        main_layout.addWidget(group_appear)

        # 윈도우 선택 버튼 (하단으로 이동)
        self.win_select_btn = QtWidgets.QPushButton("윈도우 선택 (Select Window)")
        self.win_select_btn.clicked.connect(self.sig_req_window_select.emit)
        main_layout.addWidget(self.win_select_btn)

        self.setLayout(main_layout)

    def update_config(self):
        Config.CANNY_MIN = self.canny_min_slider.value()
        Config.CANNY_MAX = self.canny_max_slider.value()
        Config.EDGE_THICKNESS = self.thickness_spin.value()
        Config.EDGE_OPACITY = self.opacity_slider.value()
        Config.AUTO_SIGMA = self.sigma_slider.value() / 100.0
        Config.REALTIME_AUTO = self.realtime_chk.isChecked()
        
        self.canny_min_label.setText(f"Canny Min: {Config.CANNY_MIN}")
        self.canny_max_label.setText(f"Canny Max: {Config.CANNY_MAX}")
        self.opacity_label.setText(f"투명도 (Opacity): {Config.EDGE_OPACITY}")
        self.sigma_label.setText(f"Sigma: {Config.AUTO_SIGMA:.2f}")
        self.valueChanged.emit()

    def update_sliders_from_config(self):
        # 재귀적인 시그널 발생 방지
        self.canny_min_slider.blockSignals(True)
        self.canny_max_slider.blockSignals(True)
        self.opacity_slider.blockSignals(True)
        self.sigma_slider.blockSignals(True)
        
        self.canny_min_slider.setValue(int(Config.CANNY_MIN))
        self.canny_max_slider.setValue(int(Config.CANNY_MAX))
        self.thickness_spin.setValue(int(Config.EDGE_THICKNESS))
        self.opacity_slider.setValue(int(Config.EDGE_OPACITY))
        self.sigma_slider.setValue(int(Config.AUTO_SIGMA * 100))
        self.realtime_chk.setChecked(Config.REALTIME_AUTO)

        self.canny_min_label.setText(f"Canny Min: {int(Config.CANNY_MIN)}")
        self.canny_max_label.setText(f"Canny Max: {int(Config.CANNY_MAX)}")
        self.opacity_label.setText(f"투명도 (Opacity): {int(Config.EDGE_OPACITY)}")
        self.sigma_label.setText(f"Sigma: {Config.AUTO_SIGMA:.2f}")
        
        self.canny_min_slider.blockSignals(False)
        self.canny_max_slider.blockSignals(False)
        self.opacity_slider.blockSignals(False)
        self.sigma_slider.blockSignals(False)


    def change_color(self):
        color = QtWidgets.QColorDialog.getColor(
            QtGui.QColor(*Config.EDGE_COLOR), self, "엣지 색상 선택"
        )
        if color.isValid():
            Config.EDGE_COLOR = (color.red(), color.green(), color.blue())
            self.valueChanged.emit()

class CaptureWorker(QtCore.QThread):
    sig_frame_ready = QtCore.pyqtSignal(QtGui.QImage)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        self.paused = False
        self.cam = None
        self.current_monitor_idx = 0
        self.region = None # (lx, ly, w, h)
        self.request_one_shot_auto = False # 1회성 자동 조절 요청 플래그
        
        # DXCam 초기화
        self._init_camera(0)

    def _init_camera(self, monitor_idx):
        try:
            # 안전하게 기존 카메라 해제
            if getattr(self, 'cam', None) is not None:
                old_cam = self.cam
                self.cam = None # 속성은 유지하되 값만 None으로 변경 (run 루프 안전)
                del old_cam
                
            self.cam = dxcam.create(output_idx=monitor_idx, output_color="BGR")
            self.current_monitor_idx = monitor_idx
        except Exception as e:
            print(f"CaptureWorker: DXCam 초기화 실패: {e}")
            self.cam = None

    def set_region_and_monitor(self, monitor_idx, region):
        self.region = region
        if self.current_monitor_idx != monitor_idx:
            # 모니터 변경 시 카메라 재초기화
            self._init_camera(monitor_idx)
            
    def trigger_auto_adjust(self):
        self.request_one_shot_auto = True

    def run(self):
        while self.running:
            if self.paused or self.cam is None or self.region is None:
                self.msleep(100)
                continue
                
            lx, ly, w, h = self.region
            if w <= 0 or h <= 0:
                self.msleep(50)
                continue
                
            # 프레임 캡처 (Grab Frame)
            try:
                # 모니터 경계 확인 (Clamp)
                if lx + w > self.cam.width: w = self.cam.width - lx
                if ly + h > self.cam.height: h = self.cam.height - ly
                
                if w <= 0 or h <= 0:
                    self.msleep(50)
                    continue

                frame = self.cam.grab(region=(lx, ly, lx+w, ly+h))
                if frame is None: 
                    continue
            except Exception as e:
                self.msleep(50)
                continue
                
            # 이미지 처리 (Process Image)
            try:
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                
                # 자동 조절 로직 (실시간 또는 1회성 요청)
                if Config.REALTIME_AUTO or self.request_one_shot_auto:
                    self._calculate_auto_threshold(gray)
                    self.request_one_shot_auto = False # 요청 처리 완료
                
                blurred = cv2.GaussianBlur(gray, (5, 5), 0)
                edges = cv2.Canny(blurred, Config.CANNY_MIN, Config.CANNY_MAX)
                
                if Config.EDGE_THICKNESS > 1:
                    kernel = np.ones((Config.EDGE_THICKNESS, Config.EDGE_THICKNESS), np.uint8)
                    edges = cv2.dilate(edges, kernel, iterations=1)
                    
                # RGBA 변환
                h_img, w_img = edges.shape
                rgba_image = np.zeros((h_img, w_img, 4), dtype=np.uint8)
                
                r, g, b = Config.EDGE_COLOR
                mask = edges > 0
                rgba_image[mask, 0] = r
                rgba_image[mask, 1] = g
                rgba_image[mask, 2] = b
                rgba_image[mask, 3] = Config.EDGE_OPACITY
                
                qt_image = QtGui.QImage(rgba_image.data, w_img, h_img, w_img * 4, QtGui.QImage.Format_RGBA8888).copy()
                self.sig_frame_ready.emit(qt_image)
                
            except Exception as e:
                print(f"이미지 처리 오류: {e}")

            # 주사율 제어 (Control refresh rate)
            if Config.REFRESH_RATE > 0:
                self.msleep(int(1000 / Config.REFRESH_RATE))
            else:
                self.msleep(16)

    def _calculate_auto_threshold(self, gray_image):
        # 메인 스레드 로직과 동일한 자동 조절 로직
        h, w = gray_image.shape
        if w > 320:
            small_gray = cv2.resize(gray_image, (0, 0), fx=0.2, fy=0.2, interpolation=cv2.INTER_NEAREST)
        else:
            small_gray = gray_image
            
        v = np.median(small_gray)
        sigma = Config.AUTO_SIGMA
        lower = int(max(0, (1.0 - sigma) * v))
        upper = int(min(255, (1.0 + sigma) * v))
        
        # 공유 설정 업데이트
        Config.CANNY_MIN = lower
        Config.CANNY_MAX = upper

    def stop(self):
        self.running = False
        self.wait()


class OverlayWindow(QtWidgets.QMainWindow):
    # 키보드 스레드에서 안전한 UI 업데이트를 위한 시그널 정의
    sig_toggle_visibility = QtCore.pyqtSignal()
    sig_toggle_interactive = QtCore.pyqtSignal()
    sig_snap_window = QtCore.pyqtSignal(int, object) # hwnd, rect

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Edge Overlay")
        
        # 윈도우 초기 크기 및 위치
        screen = QtWidgets.QApplication.primaryScreen().size()
        self.setGeometry(screen.width()//4, screen.height()//4, screen.width()//2, screen.height()//2)

        # 플래그 설정
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint |
            QtCore.Qt.WindowStaysOnTopHint |
            QtCore.Qt.Tool 
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # 캡처 루프 방지 (자기 자신 캡처 제외)
        try:
            user32 = ctypes.windll.user32
            user32.SetWindowDisplayAffinity(int(self.winId()), 0x00000011)
        except Exception as e:
            print(f"SetWindowDisplayAffinity 실패: {e}")
        
        # 워커 스레드 설정
        self.worker = CaptureWorker()
        self.worker.sig_frame_ready.connect(self.update_image_slot)
        self.worker.start()

        # 모니터 관리
        self.check_monitor_timer = QtCore.QTimer()
        self.check_monitor_timer.timeout.connect(self.update_capture_region)
        self.check_monitor_timer.start(500) # 0.5초마다 체크

        # 설정 윈도우
        self.settings_widget = SettingsWidget()
        
        # 상태 변수
        self.qt_image = None
        self.is_interactive = True
        self.is_visible = True
        self.update_mouse_input_mode()
        
        # 설정 시그널 연결
        self.settings_widget.sig_req_auto_adjust.connect(self.perform_auto_adjust)
        self.settings_widget.sig_req_window_select.connect(self.start_window_selection)
        
        # 윈도우 선택 상태
        self.is_selecting_window = False
        self.selection_timer = QtCore.QTimer()
        self.selection_timer.timeout.connect(self.process_window_selection)
        self.sig_snap_window.connect(self.finish_window_selection)
        self.highlight_rect = None
        self.pre_selection_geometry = None

        # 이동/리사이즈 관련 변수
        self.drag_pos = None
        self.resize_margin = 30
        self.resize_edge = None

        # 시그널 연결
        self.sig_toggle_visibility.connect(self.toggle_visibility_slot)
        self.sig_toggle_interactive.connect(self.toggle_interactive_slot)

        # 글로벌 단축키 등록
        keyboard.add_hotkey('ctrl+f9', lambda: self.sig_toggle_visibility.emit())
        keyboard.add_hotkey('ctrl+f10', lambda: self.sig_toggle_interactive.emit())

    def toggle_visibility_slot(self):
        self.is_visible = not self.is_visible
        if not self.is_visible:
            self.worker.paused = True
            self.qt_image = None
            self.update()
        else:
            self.worker.paused = False
            self.update_capture_region() # 깨우기

    def toggle_interactive_slot(self):
        self.is_interactive = not self.is_interactive
        self.update_mouse_input_mode()

    def update_mouse_input_mode(self):
        hwnd = self.winId().__int__()
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        
        if self.is_interactive:
            style &= ~win32con.WS_EX_TRANSPARENT
            self.settings_widget.show()
        else:
            style |= win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED
            self.settings_widget.hide()
            
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, style)
        self.update()

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            if self.is_selecting_window:
                # 선택 취소
                self.is_selecting_window = False
                self.selection_timer.stop()
                QtWidgets.QApplication.restoreOverrideCursor()
                if self.pre_selection_geometry:
                    self.setGeometry(self.pre_selection_geometry)
                self.is_interactive = True
                self.update_mouse_input_mode()
                self.worker.paused = False
            else:
                self.settings_widget.close()

    def closeEvent(self, event):
        self.worker.stop()
        keyboard.unhook_all()
        super().closeEvent(event)

    def update_image_slot(self, image):
        if self.is_visible and not self.is_selecting_window:
            self.qt_image = image
            self.update()
        
        # 자동 조절이 값을 변경했을 경우 UI 슬라이더 동기화
        if (Config.REALTIME_AUTO or self.settings_widget.auto_btn.isDown()) and self.settings_widget.isVisible():
             self.settings_widget.update_sliders_from_config()

    def update_capture_region(self):
        if not self.is_visible: return
        
        geometry = self.geometry()
        x, y, w, h = geometry.x(), geometry.y(), geometry.width(), geometry.height()
        
        # 모니터 및 안전 영역 찾기 로직
        # 1. 모니터 찾기
        cx, cy = x + w//2, y + h//2
        screen = QtWidgets.QApplication.screenAt(QtCore.QPoint(cx, cy))
        if not screen: 
            screen = QtWidgets.QApplication.primaryScreen()
            
        # 2. 모니터 지오메트리 가져오기
        screen_geo = screen.geometry()
        sx, sy, sw, sh = screen_geo.x(), screen_geo.y(), screen_geo.width(), screen_geo.height()
        
        # 3. 모니터 인덱스 결정 (휴리스틱)
        screens = QtWidgets.QApplication.screens()
        try:
            target_idx = screens.index(screen)
        except ValueError:
            target_idx = 0

        # 4. 교차 영역 계산
        ix = max(x, sx)
        iy = max(y, sy)
        iw = min(x + w, sx + sw) - ix
        ih = min(y + h, sy + sh) - iy
        
        if iw <= 0 or ih <= 0:
            return # 화면 밖

        # 5. 로컬 좌표 (DXCam용)
        lx = ix - sx
        ly = iy - sy
        
        # 워커 업데이트
        self.worker.set_region_and_monitor(target_idx, (lx, ly, iw, ih))


    def start_window_selection(self):
        self.pre_selection_geometry = self.geometry()
        self.is_selecting_window = True
        self.is_interactive = False 
        
        # 워커 일시 정지
        self.worker.paused = True
        self.qt_image = None
        
        # 모든 클릭을 잡기 위한 전체 화면 (가상 스크린)
        v_screen = QtWidgets.QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(v_screen)
        
        self.update_mouse_input_mode() # 클릭 가능하도록 설정
        
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.CrossCursor)
        self.selection_timer.start(50)

    def process_window_selection(self):
        if not self.is_selecting_window:
            self.selection_timer.stop()
            QtWidgets.QApplication.restoreOverrideCursor()
            return

        pos = win32gui.GetCursorPos() # 글로벌 좌표
        
        # 하이라이트 로직
        hwnd = win32gui.WindowFromPoint(pos)
        ancestor = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
        if ancestor: hwnd = ancestor
        
        # 자기 자신 제외
        if hwnd != int(self.winId()):
             try:
                rect = win32gui.GetWindowRect(hwnd)
                # 로컬 좌표로 변환 (현재 전체 화면 상태)
                g_geo = self.geometry()
                vx, vy = g_geo.x(), g_geo.y()
                
                l, t, r, b = rect
                self.highlight_rect = QtCore.QRect(l - vx, t - vy, r - l, b - t)
                self.update() # 다시 그리기 유도
                
             except Exception:
                 self.highlight_rect = None
        
        # 좌클릭 확인
        if win32api.GetAsyncKeyState(0x01) < 0:
            if hwnd and hwnd != int(self.winId()):
                try:
                    rect = win32gui.GetWindowRect(hwnd)
                    self.sig_snap_window.emit(hwnd, rect)
                except:
                    pass
            self.is_selecting_window = False

    def finish_window_selection(self, hwnd, rect):
        self.selection_timer.stop()
        QtWidgets.QApplication.restoreOverrideCursor()
        self.highlight_rect = None
        
        l, t, r, b = rect
        w, h = r - l, b - t
        
        # 최소 크기 보정
        if w < 100: w = 100
        if h < 100: h = 100
            
        self.setGeometry(l, t, w, h)
        self.is_interactive = True
        self.update_mouse_input_mode()
        self.worker.paused = False

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        
        if self.is_selecting_window:
            # 하이라이트 그리기
            if self.highlight_rect:
                painter.setPen(QtGui.QPen(QtCore.Qt.red, 2))
                painter.drawRect(self.highlight_rect)
            return

        if self.qt_image:
            painter.drawImage(0, 0, self.qt_image)
        
        if self.is_interactive:
            painter.setPen(QtGui.QPen(QtCore.Qt.yellow))
            painter.drawText(10, 20, "[Ctrl+F9] 보기 토글 | [Ctrl+F10] 편집 모드")
            
            # 테두리 그리기
            painter.setPen(QtGui.QPen(QtCore.Qt.yellow, 5))
            painter.drawRect(0, 0, self.width()-1, self.height()-1)

    def get_resize_edge(self, pos):
        rect = self.rect()
        edge = 0
        x, y = pos.x(), pos.y()
        w, h = rect.width(), rect.height()
        m = self.resize_margin

        if x < m: edge |= 1  # 좌
        if x > w - m: edge |= 2  # 우
        if y < m: edge |= 4  # 상
        if y > h - m: edge |= 8  # 하
        return edge

    def mousePressEvent(self, event):
        if not self.is_interactive: return
        if event.button() == QtCore.Qt.LeftButton:
            self.resize_edge = self.get_resize_edge(event.pos())
            if self.resize_edge == 0:
                self.drag_pos = event.globalPos()
            else:
                self.drag_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if not self.is_interactive: return
        
        edge = self.get_resize_edge(event.pos())
        if edge == 0:
            self.setCursor(QtCore.Qt.ArrowCursor)
        elif edge in (1, 2):
            self.setCursor(QtCore.Qt.SizeHorCursor)
        elif edge in (4, 8): 
            self.setCursor(QtCore.Qt.SizeVerCursor)
        else:
            self.setCursor(QtCore.Qt.SizeFDiagCursor if edge in (5, 10) else QtCore.Qt.SizeBDiagCursor)

        if event.buttons() & QtCore.Qt.LeftButton:
            if self.resize_edge == 0 and self.drag_pos:
                delta = event.globalPos() - self.drag_pos
                self.move(self.pos() + delta)
                self.drag_pos = event.globalPos()
            elif self.resize_edge != 0 and self.drag_pos:
                delta = event.globalPos() - self.drag_pos
                geo = self.geometry()
                
                if self.resize_edge & 1: geo.setLeft(geo.left() + delta.x())
                if self.resize_edge & 2: geo.setRight(geo.right() + delta.x())
                if self.resize_edge & 4: geo.setTop(geo.top() + delta.y())
                if self.resize_edge & 8: geo.setBottom(geo.bottom() + delta.y())
                
                # 최소 크기 제한
                if geo.width() < 100: geo.setWidth(100)
                if geo.height() < 100: geo.setHeight(100)
                
                self.setGeometry(geo)
                self.drag_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.drag_pos = None
        self.resize_edge = None

    def perform_auto_adjust(self):
        # 워커에게 1회성 자동 조절 요청
        self.worker.trigger_auto_adjust()

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = OverlayWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
