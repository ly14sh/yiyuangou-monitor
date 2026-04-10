"""
一元购 · 库存监控系统 v4.0
功能：监控 vtravel.link2shops.com 一元购活动库存，有货时播放报警并自动停止监控
支持：单商品选择、WAV 音乐 / TTS 语音播报、系统托盘
"""

import sys
import os
import json
import subprocess
import struct
import wave
import math
import winsound
import requests
from datetime import datetime, time as dtime

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QTextEdit, QGroupBox, QFormLayout,
    QSystemTrayIcon, QMenu, QMessageBox,
    QSpinBox, QStatusBar, QComboBox, QTimeEdit
)
from PySide6.QtCore import QTimer, Qt, QTime
from PySide6.QtGui import QAction, QColor, QBrush, QPainter, QPixmap, QIcon


# ─── 工作目录 & 资源 ───────────────────────────────────────────
WORKSPACE    = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE  = os.path.join(WORKSPACE, "config.json")
WAV_PATH     = os.path.join(WORKSPACE, "alert.wav")

# ─── 全局常量 ───────────────────────────────────────────────────
# 音乐播放参数改为从配置读取，不再硬编码

# ─── API 配置 ───────────────────────────────────────────────────
API_URL = "https://vtravel.link2shops.com/vfuliApi/api/client/ypJyActivity/goodsDetail"
API_HEADERS = {
    "Host": "vtravel.link2shops.com",
    "Connection": "keep-alive",
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Origin": "https://vtravel.link2shops.com",
    "X-Requested-With": "com.tencent.mm",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-Mode": "cors",
    "Referer": "https://vtravel.link2shops.com/yiyuan/",
    "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
}

# ─── 商品目录 ────────────────────────────────────────────────────
PRODUCTS = [
    {"id": "bilibili_monthly", "label": "B站大会员月卡",
     "activityId": "08c02ec083ea4d159b95118ba332a614",
     "goodsId":    "bfd7c35c050f4181b4ae08eb9c19f0a5",
     "platformTp": "T0060", "price": 30,
     "channelId":  "73b34182aaed4559b56e5504801f557b"},
    {"id": "youku_monthly", "label": "优酷VIP会员(月卡)",
     "activityId": "08c02ec083ea4d159b95118ba332a614",
     "goodsId":    "bf3a97780faf48118a301a151d726ae0",
     "platformTp": "T0060", "price": 30,
     "channelId":  "73b34182aaed4559b56e5504801f557b"},
    {"id": "feike_2000", "label": "飞客2000里程券",
     "activityId": "a22a226287cc4959abeca00201015478",
     "goodsId":    "8b3af0aa046c4e5c82420699c4bf9e03",
     "platformTp": "T0002", "price": 20,
     "channelId":  "73b34182aaed4559b56e5504801f557b"},
    {"id": "starbucks_43", "label": "星巴克43元星礼包",
     "activityId": "08c02ec083ea4d159b95118ba332a614",
     "goodsId":    "69041c2f1302417f938b1ea83bb2534b",
     "platformTp": "T0060", "price": 43,
     "channelId":  "73b34182aaed4559b56e5504801f557b"},
    {"id": "iqiyi_gold", "label": "爱奇艺黄金会员(月卡)",
     "activityId": "08c02ec083ea4d159b95118ba332a614",
     "goodsId":    "f7e22be73c6e419e9cb125cf1e8dad04",
     "platformTp": "T0060", "price": None,
     "channelId":  "73b34182aaed4559b56e5504801f557b"},
]


# ─── 托盘图标 ────────────────────────────────────────────────────
TRAY_GREEN, TRAY_RED, TRAY_YELLOW = 0, 1, 2

def _create_tray_icons():
    """预创建三个托盘图标，避免运行时重复绘制"""
    icons = {}
    size = 32
    colors = {TRAY_GREEN: "#27AE60", TRAY_RED: "#E74C3C", TRAY_YELLOW: "#F39C12"}
    for mode, hex_color in colors.items():
        pm = QPixmap(size, size)
        pm.fill(Qt.transparent)
        color = QColor(hex_color)
        painter = QPainter(pm)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(color))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(4, 4, 24, 24)
        painter.end()
        icons[mode] = QIcon(pm)
    return icons

# 图标缓存（延迟初始化，等 QApplication 创建后再生成）
_TRAY_ICONS = None

def make_tray_icon(color_mode):
    global _TRAY_ICONS
    if _TRAY_ICONS is None:
        _TRAY_ICONS = _create_tray_icons()
    return _TRAY_ICONS.get(color_mode, _TRAY_ICONS[TRAY_RED])


# ─── 配置读写 ───────────────────────────────────────────────────
def load_config():
    defaults = {
        "start_hour": 8, "start_min": 0,
        "end_hour": 23, "end_min": 59,
        "interval": 30,
        "alert_text": "一元购有库存啦，快去下单！",
        "selected_id": "bilibili_monthly",
        "sound_mode": "wav",   # "wav" | "tts"
        "music_gap_sec": 2,    # 播放间隔：1/2/3 秒
        "music_duration_min": 3,  # 连续时长：1/2/3 分钟
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return {**defaults, **json.load(f)}
        except Exception:
            return defaults
    return defaults

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ─── 主窗口 ─────────────────────────────────────────────────────
class MonitorWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.cfg           = load_config()
        self.monitoring    = False
        self.music_playing = False
        self._quitting     = False
        self._tts_proc     = None  # 当前TTS进程

        # 定时器
        self.check_timer   = QTimer()
        self.check_timer.timeout.connect(self._do_check)
        self.music_timer   = QTimer()
        self.music_timer.setSingleShot(True)
        self.music_timer.timeout.connect(self._music_do_play)
        self.timeout_timer = QTimer()
        self.timeout_timer.timeout.connect(self._music_stop)

        # 托盘
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(make_tray_icon(TRAY_RED))
        self.tray.setToolTip("一元购 · 库存监控")
        self.tray.show()
        tray_menu = QMenu()
        act_show = QAction("显示窗口", self)
        act_show.triggered.connect(self.showNormal)
        self.act_toggle = QAction("开始监控", self)
        self.act_toggle.triggered.connect(self._toggle_monitor)
        act_quit = QAction("退出", self)
        act_quit.triggered.connect(self.quit_app)
        tray_menu.addActions([act_show, self.act_toggle, act_quit])
        self.tray.setContextMenu(tray_menu)
        self.tray.activated.connect(
            lambda r: self.showNormal() if r == QSystemTrayIcon.DoubleClick else None
        )

        self._build_ui()
        self._apply_cfg_to_ui()

    # ─── 声音播报 ──────────────────────────────────────────────
    def _music_do_play(self):
        if not self.music_playing:
            return
        mode = self.cfg.get("sound_mode", "wav")
        try:
            if mode == "wav":
                # 使用用户指定的WAV路径，优先从配置读取
                wav_file = self.cfg.get("wav_path", WAV_PATH)
                if wav_file and os.path.exists(wav_file):
                    winsound.PlaySound(wav_file, winsound.SND_FILENAME | winsound.SND_ASYNC)
                else:
                    # WAV文件不存在，自动切换到TTS并提示
                    self._log("[提示] WAV文件不存在，已自动切换到TTS语音播报")
                    mode = "tts"
            if mode == "tts":
                # TTS：System.Speech（非阻塞执行，先终止旧进程避免重叠）
                if self._tts_proc and self._tts_proc.poll() is None:
                    try:
                        self._tts_proc.terminate()
                    except Exception:
                        pass
                text = self.edit_sound.text().strip() or "一元购有库存啦，快去下单！"
                escaped = text.replace("'", "''")
                ps = (
                    "Add-Type -AssemblyName System.Speech; "
                    "(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('" + escaped + "')"
                )
                self._tts_proc = subprocess.Popen(
                    ["powershell", "-NoProfile", "-Command", ps],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    creationflags=0x08000000
                )
        except Exception as e:
            self._log(f"[警告] 声音播放失败: {e}")
        if self.music_playing:
            # 直接从 UI 读取当前设置的间隔（秒转毫秒）
            gap_sec = self.combo_gap.currentData() or 2
            self.music_timer.start(gap_sec * 1000)

    def _music_play(self):
        self.music_playing = True
        self._music_do_play()
        # 直接从 UI 读取当前设置的连续时长（分钟转毫秒）
        dur_min = self.combo_duration.currentData() or 3
        self.timeout_timer.start(dur_min * 60 * 1000)
        self._update_sound_btn()

    def _music_stop(self):
        self.music_playing = False
        self.music_timer.stop()
        self.timeout_timer.stop()
        try:
            winsound.PlaySound(None, 0)  # 停止正在播放的 WAV
        except Exception:
            pass
        # 终止TTS进程
        if self._tts_proc and self._tts_proc.poll() is None:
            try:
                self._tts_proc.terminate()
            except Exception:
                pass
            self._tts_proc = None
        self._update_sound_btn()

    def _update_sound_btn(self):
        if self.music_playing:
            label = "⏹ 停止声音"
            color = "#E67E22"
        else:
            label = "🔊 测试声音"
            color = "#8E44AD"
        self.btn_sound.setText(label)
        self.btn_sound.setStyleSheet(
            f"background:{color}; color:white; font-size:13px; font-weight:bold; "
            "padding:6px 14px; border-radius:4px;"
        )

    # ─── UI ────────────────────────────────────────────────────
    def _build_ui(self):
        lay = QVBoxLayout(self)

        # ── 状态区 ─────────────────────────────────────────
        status_box = QGroupBox("当前状态")
        s_layout = QFormLayout()
        self.lbl_status  = QLabel("● 已停止")
        self.lbl_status.setStyleSheet("color:#E74C3C; font-size:15px; font-weight:bold;")
        self.lbl_summary = QLabel("—")
        self.lbl_next    = QLabel("—")
        s_layout.addRow("运行状态：", self.lbl_status)
        s_layout.addRow("库存状态：", self.lbl_summary)
        s_layout.addRow("下次检测：", self.lbl_next)
        status_box.setLayout(s_layout)
        lay.addWidget(status_box)

        # ── 商品选择（单选下拉）────────────────────────────
        prod_box = QGroupBox("选择监控商品")
        p_layout = QFormLayout()
        self.combo_product = QComboBox()
        for p in PRODUCTS:
            price_str = f"¥{p['price']}元" if p["price"] else ""
            self.combo_product.addItem(f"{p['label']}  {price_str}", userData=p["id"])
        p_layout.addRow("监控商品：", self.combo_product)
        prod_box.setLayout(p_layout)
        lay.addWidget(prod_box)

        # ── 设置区 ─────────────────────────────────────────
        set_box = QGroupBox("监控设置")
        f_layout = QFormLayout()

        # 时间段
        time_hbox = QHBoxLayout()
        self.time_start = QTimeEdit()
        self.time_start.setDisplayFormat("HH:mm")
        self.time_end = QTimeEdit()
        self.time_end.setDisplayFormat("HH:mm")
        time_hbox.addWidget(self.time_start)
        time_hbox.addWidget(QLabel("  —  "))
        time_hbox.addWidget(self.time_end)
        time_hbox.addStretch()
        f_layout.addRow("监控时段：", time_hbox)

        # 检测间隔
        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(5, 300)
        self.spin_interval.setSuffix(" 秒")
        f_layout.addRow("检测间隔：", self.spin_interval)

        # 声音模式选择
        self.combo_sound_mode = QComboBox()
        self.combo_sound_mode.addItems(["🔔 警报音乐（WAV）", "🔉 TTS 语音播报"])
        self.combo_sound_mode.setCurrentIndex(0)
        self.combo_sound_mode.currentIndexChanged.connect(self._on_sound_mode_changed)
        f_layout.addRow("声音模式：", self.combo_sound_mode)

        # WAV文件路径选择（WAV模式显示）
        wav_hbox = QHBoxLayout()
        self.edit_wav_path = QLineEdit()
        self.edit_wav_path.setPlaceholderText("默认：程序目录下的 alert.wav")
        self.edit_wav_path.setText(WAV_PATH if os.path.exists(WAV_PATH) else "")
        self.edit_wav_path.textChanged.connect(lambda t: self.cfg.update({"wav_path": t}))
        self.btn_browse = QPushButton("浏览")
        self.btn_browse.setStyleSheet("padding:2px 8px;")
        self.btn_browse.clicked.connect(self._browse_wav)
        wav_hbox.addWidget(self.edit_wav_path)
        wav_hbox.addWidget(self.btn_browse)
        self.wav_row_widget = QWidget()
        self.wav_row_widget.setLayout(wav_hbox)
        self.wav_row_label = QLabel("WAV 文件：")
        f_layout.addRow(self.wav_row_label, self.wav_row_widget)

        # TTS 语音内容（TTS模式显示）
        self.edit_sound = QLineEdit()
        self.edit_sound.setPlaceholderText("有货时的语音播报内容...")
        self.tts_row_label = QLabel("TTS 内容：")
        f_layout.addRow(self.tts_row_label, self.edit_sound)
        # 默认WAV模式：隐藏TTS行
        self.edit_sound.setVisible(False)
        self.tts_row_label.setVisible(False)

        # 播报间隔
        self.combo_gap = QComboBox()
        for i in range(1, 10):
            self.combo_gap.addItem(f"{i} 秒", i)
        f_layout.addRow("播报间隔：", self.combo_gap)

        # 连续时长
        self.combo_duration = QComboBox()
        for i in range(1, 6):
            self.combo_duration.addItem(f"{i} 分钟", i)
        f_layout.addRow("连续时长：", self.combo_duration)

        # 按钮行
        self.btn_toggle = QPushButton("▶ 开始监控")
        self.btn_toggle.setStyleSheet(
            "background:#27AE60; color:white; font-size:14px; font-weight:bold; "
            "padding:6px 20px; border-radius:4px;"
        )
        self.btn_toggle.clicked.connect(self._toggle_monitor)

        self.btn_sound = QPushButton("🔊 测试声音")
        self.btn_sound.setStyleSheet(
            "background:#8E44AD; color:white; font-size:13px; font-weight:bold; "
            "padding:6px 14px; border-radius:4px;"
        )
        self.btn_sound.clicked.connect(self._on_sound_btn)

        hbox_btn = QHBoxLayout()
        hbox_btn.addWidget(self.btn_toggle)
        hbox_btn.addWidget(self.btn_sound)
        hbox_btn.addStretch()
        f_layout.addRow("", hbox_btn)

        set_box.setLayout(f_layout)
        lay.addWidget(set_box)

        # ── 日志区 ─────────────────────────────────────────
        log_box = QGroupBox("检测日志")
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMaximumHeight(130)
        log_box.setLayout(QVBoxLayout())
        log_box.layout().addWidget(self.log_edit)
        lay.addWidget(log_box)

        # ── 状态栏 ─────────────────────────────────────────
        self.status_bar = QStatusBar()
        lay.addWidget(self.status_bar)
        self.status_bar.showMessage("就绪")

    def _on_sound_mode_changed(self, index):
        mode = "wav" if index == 0 else "tts"
        self.cfg["sound_mode"] = mode
        is_wav = (mode == "wav")
        self.wav_row_widget.setVisible(is_wav)
        self.wav_row_label.setVisible(is_wav)
        self.edit_sound.setVisible(not is_wav)
        self.tts_row_label.setVisible(not is_wav)
        self._log(f"声音模式切换为：{'WAV 警报音' if is_wav else 'TTS 语音播报'}")

    def _browse_wav(self):
        from PySide6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getOpenFileName(
            self, "选择WAV音频文件", "",
            "音频文件 (*.wav);;所有文件 (*.*)"
        )
        if path:
            self.edit_wav_path.setText(path)
            self.cfg["wav_path"] = path

    def _find_tts_label(self):
        # 通过 layout 找 TTS 行标签
        for label in self.findChildren(QLabel):
            if "TTS 内容" in label.text():
                return label
        return None

    def _on_sound_btn(self):
        if self.music_playing:
            self._music_stop()
            self._log("已停止声音")
        else:
            # 临时启动一次测试（不循环，自动停止）
            self._music_play()
            self._log("正在测试声音...")
            QTimer.singleShot(5000, self._music_stop)

    def _apply_cfg_to_ui(self):
        self.time_start.setTime(
            QTime(self.cfg["start_hour"], self.cfg["start_min"])
        )
        self.time_end.setTime(
            QTime(self.cfg["end_hour"], self.cfg["end_min"])
        )
        self.spin_interval.setValue(self.cfg.get("interval", 30))
        self.edit_sound.setText(self.cfg.get("alert_text", "一元购有库存啦，快去下单！"))

        # 商品选中
        sel_id = self.cfg.get("selected_id", "bilibili_monthly")
        for i in range(self.combo_product.count()):
            if self.combo_product.itemData(i) == sel_id:
                self.combo_product.setCurrentIndex(i)
                break

        # 声音模式
        mode = self.cfg.get("sound_mode", "wav")
        self.combo_sound_mode.setCurrentIndex(0 if mode == "wav" else 1)
        is_wav = (mode == "wav")
        self.wav_row_widget.setVisible(is_wav)
        self.wav_row_label.setVisible(is_wav)
        self.edit_sound.setVisible(not is_wav)
        self.tts_row_label.setVisible(not is_wav)
        # WAV路径
        wav_path = self.cfg.get("wav_path", WAV_PATH)
        self.edit_wav_path.setText(wav_path if os.path.exists(wav_path) else "")

        # 播报间隔 & 连续时长
        gap = self.cfg.get("music_gap_sec", 2)
        for i in range(self.combo_gap.count()):
            if self.combo_gap.itemData(i) == gap:
                self.combo_gap.setCurrentIndex(i)
                break
        dur = self.cfg.get("music_duration_min", 3)
        for i in range(self.combo_duration.count()):
            if self.combo_duration.itemData(i) == dur:
                self.combo_duration.setCurrentIndex(i)
                break

    def _save_cfg_from_ui(self):
        self.cfg["start_hour"] = self.time_start.time().hour()
        self.cfg["start_min"]  = self.time_start.time().minute()
        self.cfg["end_hour"]   = self.time_end.time().hour()
        self.cfg["end_min"]    = self.time_end.time().minute()
        self.cfg["interval"]   = self.spin_interval.value()
        self.cfg["alert_text"] = self.edit_sound.text().strip() or "一元购有库存啦，快去下单！"
        self.cfg["selected_id"] = self.combo_product.currentData()
        self.cfg["music_gap_sec"] = self.combo_gap.currentData()
        self.cfg["music_duration_min"] = self.combo_duration.currentData()
        save_config(self.cfg)

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_edit.append(f"[{ts}] {msg}")
        self.status_bar.showMessage(msg)

    def _selected_product(self):
        sel_id = self.combo_product.currentData()
        return next((p for p in PRODUCTS if p["id"] == sel_id), None)

    # ─── 库存检测 ──────────────────────────────────────────────
    def _check_stock(self, product):
        payload = {
            "activityId": product["activityId"],
            "goodsId":    product["goodsId"],
            "channelId":  product["channelId"],
            "platformTp": product["platformTp"],
        }
        try:
            resp = requests.post(API_URL, json=payload, headers=API_HEADERS, timeout=10)
            resp.encoding = "utf-8"
            data = resp.json()
            if data.get("code") != 0:
                return None, data.get("msg", "未知错误")
            goods = data.get("goodsMap") or data.get("data", {}).get("goodsMap") or {}
            stock   = goods.get("stock", 0)
            status  = goods.get("stockStatus", "2")
            name    = goods.get("name", product["label"])
            price   = goods.get("purchasePrice") or product.get("price") or "?"
            return name, stock, status, price
        except Exception as e:
            return None, f"请求异常: {e}"

    def _is_in_time_window(self):
        now    = datetime.now()
        cur    = dtime(now.hour, now.minute)
        start  = dtime(self.time_start.time().hour(), self.time_start.time().minute())
        end    = dtime(self.time_end.time().hour(),   self.time_end.time().minute())
        if start <= end:
            return start <= cur <= end
        return cur >= start or cur <= end

    def _do_check(self):
        if not self._is_in_time_window():
            self.lbl_summary.setText("⏸ 不在监控时段")
            self.lbl_summary.setStyleSheet("color:#7f8c8d; font-weight:bold;")
            return

        product = self._selected_product()
        if not product:
            self.lbl_summary.setText("⚠ 未选择商品")
            self.lbl_summary.setStyleSheet("color:#F39C12; font-weight:bold;")
            self._log("未选择商品")
            return

        name, stock, status, price = self._check_stock(product)

        if name is None:
            self.lbl_summary.setText(f"⚠ 查询失败: {stock}")
            self.lbl_summary.setStyleSheet("color:#F39C12; font-weight:bold;")
            self._log(f"[{product['label']}] {stock}")
        elif status in ("0", "1") or (isinstance(stock, int) and stock > 0):
            self.lbl_summary.setText(f"🎉 有货！库存:{stock}  ¥{price}")
            self.lbl_summary.setStyleSheet("color:#27AE60; font-size:14px; font-weight:bold;")
            self._log(f"🟢 [{name}] 有货！库存:{stock}  ¥{price}")
            # 有货 → 先弹窗 → 再播放声音（避免TTS阻塞弹窗显示）
            self._on_stock_found(product)
            self._music_play()
            return
        else:
            self.lbl_summary.setText(f"❌ 售罄 | 库存:{stock}")
            self.lbl_summary.setStyleSheet("color:#E74C3C;")
            self._log(f"❌ [{name}] 售罄 | 库存:{stock}")

        nxt = datetime.now().timestamp() + self.spin_interval.value()
        self.lbl_next.setText(datetime.fromtimestamp(nxt).strftime("%H:%M:%S"))

    def _on_stock_found(self, product):
        # 先停止监控（但保留语音播报），避免重复弹窗
        self._stop_monitor(stop_music=False)
        
        dur_min = self.combo_duration.currentData() or 3
        price_str = f"¥{product['price']}元" if product["price"] else ""
        msg = QMessageBox(self)
        msg.setWindowTitle("🎉 有货啦！")
        msg.setText(f"{product['label']} {price_str} 有库存了！\n\n"
                    f"快去下单：https://vtravel.link2shops.com/yiyuan/\n\n"
                    f"（声音播报中，约 {dur_min} 分钟后自动停止）")
        msg.setIcon(QMessageBox.Information)
        # 设置为非模态，避免锁定整个屏幕区域
        msg.setWindowModality(Qt.NonModal)
        msg.show()
        # 强制置顶并激活
        msg.raise_()
        msg.activateWindow()

    def _start_monitor(self):
        product = self._selected_product()
        if not product:
            QMessageBox.warning(self, "未选择商品", "请先选择要监控的商品！")
            return
        self._save_cfg_from_ui()
        self.monitoring = True
        self.check_timer.start(self.spin_interval.value() * 1000)
        self.lbl_status.setText("● 监控中")
        self.lbl_status.setStyleSheet("color:#27AE60; font-size:15px; font-weight:bold;")
        self.tray.setIcon(make_tray_icon(TRAY_YELLOW))
        self.act_toggle.setText("停止监控")
        self.btn_toggle.setText("■ 停止监控")
        self.btn_toggle.setStyleSheet(
            "background:#E74C3C; color:white; font-size:14px; font-weight:bold; "
            "padding:6px 20px; border-radius:4px;"
        )
        self._log(f"监控已启动：{product['label']}")
        self._do_check()  # 立即查一次

    def _stop_monitor(self, stop_music=True):
        self.monitoring = False
        self.check_timer.stop()
        if stop_music:
            self.music_timer.stop()
            self.timeout_timer.stop()
            try:
                winsound.PlaySound(None, 0)
            except Exception:
                pass
        self.lbl_status.setText("● 已停止")
        self.lbl_status.setStyleSheet("color:#E74C3C; font-size:15px; font-weight:bold;")
        self.lbl_summary.setText("—")
        self.lbl_next.setText("—")
        self.tray.setIcon(make_tray_icon(TRAY_RED))
        self.act_toggle.setText("开始监控")
        self.btn_toggle.setText("▶ 开始监控")
        self.btn_toggle.setStyleSheet(
            "background:#27AE60; color:white; font-size:14px; font-weight:bold; "
            "padding:6px 20px; border-radius:4px;"
        )
        self._log("监控已停止")
        self._update_sound_btn()

    def _toggle_monitor(self):
        if self.monitoring:
            self._stop_monitor()
        else:
            self._start_monitor()

    def quit_app(self):
        self._quitting = True
        try:
            winsound.PlaySound(None, 0)
        except Exception:
            pass
        # 终止TTS进程
        if self._tts_proc and self._tts_proc.poll() is None:
            try:
                self._tts_proc.terminate()
            except Exception:
                pass
        QApplication.quit()

    def closeEvent(self, event):
        if self._quitting:
            event.accept()
            return
        # 弹出选择对话框：关闭还是最小化到托盘
        reply = QMessageBox.question(
            self, "退出确认",
            "您想要：\n\n"
            "• 是 → 最小化到系统托盘（后台继续运行）\n"
            "• 否 → 完全退出程序",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            # 最小化到托盘
            event.ignore()
            self.hide()
            self.tray.showMessage(
                "一元购监控", "程序已最小化到托盘，后台继续监控", QSystemTrayIcon.Information, 3000
            )
        elif reply == QMessageBox.No:
            # 完全退出
            self._quitting = True
            self.tray.hide()
            event.accept()
        else:
            # 取消，不做任何事
            event.ignore()


# ─── 入口 ───────────────────────────────────────────────────────
if __name__ == "__main__":
    # 生成 alert.wav（如果不存在）
    if not os.path.exists(WAV_PATH):
        SAMPLE_RATE = 16000
        DURATION    = 0.8
        FREQ        = 880
        n = int(SAMPLE_RATE * DURATION)
        with wave.open(WAV_PATH, "w") as wav:
            wav.setnchannels(1)
            wav.setsampwidth(2)
            wav.setframerate(SAMPLE_RATE)
            for i in range(n):
                t = i / SAMPLE_RATE
                env = min(1.0, min(t / 0.05, (DURATION - t) / 0.05))
                val = int(env * 20000 * math.sin(2 * math.pi * FREQ * t))
                wav.writeframes(struct.pack("<h", val))

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    w = MonitorWindow()
    w.show()
    sys.exit(app.exec())
