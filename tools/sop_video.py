"""
把操作 SOP 做成影片（H.265 / HEVC 的 MP4）。

流程
  1. preview.py 用 python-pptx 反解正式的 .pptx，產出帶座標的 HTML
     —— 直接讀 pptx，影片不會和投影片走樣
  2. headless Chrome 一次截下整頁（每張投影片 1920x1080），PIL 切成 33 張
  3. 旁白：Windows SAPI 的 zh-TW 語音（Microsoft Hanhan），逐張轉 WAV
  4. ffmpeg（imageio-ffmpeg 自帶，含 libx265）把「一張圖 + 一段旁白」
     串成影片，每張的長度就是該段旁白的長度

用法
    python tools/sop_video.py            # 全部重做
    python tools/sop_video.py --no-tts   # 不配旁白，用固定秒數
"""
import base64
import json
import re
import shutil
import subprocess
import sys
import time
import wave
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = Path(r"F:\AI_BOM")
SOP = HERE / "docs" / "sop"
WORK = HERE / "docs" / "_video"
OUT_MP4 = HERE / "docs" / "BOM_Matcher_v0.10_操作SOP.mp4"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
W, H = 1920, 1080
PORT = 8830
CDP = 9333
NO_TTS = "--no-tts" in sys.argv
REUSE = "--reuse-tts" in sys.argv   # 旁白沒改時省掉重新合成

# 投影片沒有旁白時的秒數（封面、章節分隔頁之類）
MIN_SEC, PAD_SEC = 3.0, 0.8


# ── 1. 旁白文字：直接取 content.js 的結構化內容 ─────────────────────────
def narration():
    """
    回傳 [[句1, 句2, ...], ...]，一張投影片一組。

    一張一句太長（旁白十幾秒、字幕一整面），所以用標題／每條說明／注意
    各自當一句 —— 剛好也是字幕最自然的切點。
    """
    js = subprocess.run(
        ["node", "-e",
         "const c=require('F:/AI_BOM/docs/sop/content.js');"
         "console.log(JSON.stringify(c.SLIDES))"],
        capture_output=True, text=True, encoding="utf-8", check=True).stdout
    slides = json.loads(js)

    def clean(t):
        t = re.sub(r"\*\*(.+?)\*\*", r"\1", str(t))
        t = t.replace("\\n", "，").replace(chr(10), "，")
        t = re.sub(r"[\U0001F300-\U0001FAFF\u2190-\u27BF\uFE0F]", "", t)
        return re.sub(r"\s+", " ", t).strip(" ，。")

    out = []
    for s_ in slides:
        title = clean(s_.get("title") or s_.get("sub") or "")
        cues = [title] if title else []
        if s_["kind"] == "cover":
            cues += [clean(x) for x in s_.get("meta", [])]
        elif s_["kind"] != "divider":
            cues += [clean(x) for x in s_.get("bullets", [])]
            for st in s_.get("steps", []):
                cues.append(clean(st["t"]) + "，" + clean(st["d"]))
            if s_.get("warn"):
                cues.append("注意：" + clean(s_["warn"]))
        out.append([c for c in cues if c])

    out.insert(1, ["目錄", "以下是這份操作說明的目錄"])
    return out


def srt_time(t):
    h, r = divmod(t, 3600)
    m, sec = divmod(r, 60)
    return f"{int(h):02d}:{int(m):02d}:{sec:06.3f}".replace(".", ",")


def wrap(t, n=30):
    """
    字幕斷行。優先在標點或空白處斷 —— 單純數到第 n 個字就切，
    會把「2026-10-21」這種切成「2026-10-2」+「1」。
    """
    BREAK = "，。、：；！？ "
    lines, cur = [], ""
    for ch in t:
        cur += ch
        if len(cur) >= n and (ch in BREAK or len(cur) >= n + 8):
            lines.append(cur.strip()); cur = ""
    if cur.strip():
        lines.append(cur.strip())
    return "\n".join(lines[:3])


# ── 2. 投影片轉圖 ───────────────────────────────────────────────────────
def render_frames(n_slides):
    """
    逐張截圖。

    不要用「整頁截圖再平均切」—— preview.py 每張投影片上面還有一個
    「第 N 頁」標籤（div.no），平均切一定錯位，而且錯多少要靠猜。
    改用 CDP 的 Page.captureScreenshot + clip：直接拿每個 .slide 的實際
    座標來裁，版面怎麼變都不影響。
    """
    import asyncio
    import websockets
    import urllib.request

    WORK.mkdir(parents=True, exist_ok=True)
    print("  重建 HTML（反解 pptx）…")
    subprocess.run([sys.executable, str(SOP / "preview.py")],
                   check=True, capture_output=True)
    html = next((HERE / "docs").glob("BOM_Matcher_v*_操作SOP.html"))
    shutil.copy(html, WORK / "frames.html")
    html.unlink()          # HTML 版已停止維護，這裡只是中繼產物

    frames = []

    async def grab():
        srv = subprocess.Popen(
            [sys.executable, "-m", "http.server", str(PORT), "--bind", "127.0.0.1"],
            cwd=WORK, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        chrome = subprocess.Popen([
            CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
            "--hide-scrollbars", "--force-device-scale-factor=1",
            f"--window-size={W // 2},{H // 2}",
            f"--remote-debugging-port={CDP}",
            f"--user-data-dir={WORK / '_chrome'}",
            f"http://127.0.0.1:{PORT}/frames.html"])
        try:
            ws_url = None
            for _ in range(60):
                time.sleep(0.5)
                try:
                    with urllib.request.urlopen(f"http://127.0.0.1:{CDP}/json") as r:
                        tg = [t for t in json.load(r)
                              if t["type"] == "page" and "frames.html" in t.get("url", "")]
                    if tg:
                        ws_url = tg[0]["webSocketDebuggerUrl"]
                        break
                except Exception:
                    pass
            if not ws_url:
                raise SystemExit("接不上 Chrome 的除錯埠")

            async with websockets.connect(ws_url, max_size=256 * 1024 * 1024) as ws:
                seq = 0

                async def cmd(method, params=None):
                    nonlocal seq
                    seq += 1
                    await ws.send(json.dumps({"id": seq, "method": method,
                                              "params": params or {}}))
                    while True:
                        m = json.loads(await ws.recv())
                        if m.get("id") == seq:
                            if "error" in m:
                                raise SystemExit(f"{method}: {m['error']}")
                            return m.get("result", {})

                await cmd("Page.enable")
                await asyncio.sleep(6)          # 等圖片載完

                # preview.py 會把爆框的元素標成紅框（.bad）。
                # 那是給人看的檢查標記，不可以燒進影片 —— 先確認沒有問題，
                # 再把標記關掉。這一關也補回了停掉 HTML 版 SOP 之後
                # 失去的「版面爆框檢查」。
                rep0 = await cmd("Runtime.evaluate", {
                    "expression": "JSON.stringify(window.__report||{})",
                    "returnByValue": True})
                issues = json.loads(rep0["result"]["value"] or "{}").get("issues", [])
                if issues and "--ignore-layout" not in sys.argv:
                    print("  版面有問題，先修好再做影片：")
                    for x in issues:
                        print("   ", x)
                    raise SystemExit(
                        "（確定要照做可以加 --ignore-layout）")
                await cmd("Runtime.evaluate", {"expression":
                    "document.querySelectorAll('.bad')"
                    ".forEach(e=>e.classList.remove('bad'))"})
                r = await cmd("Runtime.evaluate", {
                    "expression":
                        "JSON.stringify([...document.querySelectorAll('.slide')]"
                        ".map(e=>{const b=e.getBoundingClientRect();"
                        "return {x:b.x+scrollX,y:b.y+scrollY,w:b.width,h:b.height}}))",
                    "returnByValue": True})
                boxes = json.loads(r["result"]["value"])
                if len(boxes) != n_slides:
                    raise SystemExit(f"HTML 裡有 {len(boxes)} 張投影片，"
                                     f"但 content.js 算出 {n_slides} 張")

                for i, b in enumerate(boxes, 1):
                    shot = await cmd("Page.captureScreenshot", {
                        "format": "png",
                        "captureBeyondViewport": True,
                        "clip": {"x": b["x"], "y": b["y"], "width": b["w"],
                                 "height": b["h"], "scale": W / b["w"]}})
                    f = WORK / f"f{i:03d}.png"
                    f.write_bytes(base64.b64decode(shot["data"]))
                    frames.append(f)
        finally:
            chrome.terminate()
            srv.terminate()

    asyncio.run(grab())

    from PIL import Image
    bad = [f.name for f in frames if Image.open(f).size != (W, H)]
    if bad:
        raise SystemExit(f"這些張的尺寸不是 {W}x{H}：{bad[:5]}")
    print(f"  {len(frames)} 張，每張 {W}x{H}")
    return frames


# ── 3. 旁白轉 WAV（Windows SAPI，zh-TW）────────────────────────────────
def tts(cues_per_slide):
    """逐句合成。回傳 [[wav, ...], ...]，和 cues_per_slide 同結構。"""
    flat = [{"i": i + 1, "t": t}
            for i, t in enumerate(c for g in cues_per_slide for c in g)]
    (WORK / "tts.json").write_text(json.dumps(flat, ensure_ascii=False),
                                   encoding="utf-8")
    ps = WORK / "tts.ps1"
    ps.write_text(r"""
Add-Type -AssemblyName System.Speech
$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
$zh = $s.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Culture.Name -like 'zh*' } |
      Select-Object -First 1
if ($zh) { $s.SelectVoice($zh.VoiceInfo.Name) }
$s.Rate = 0
$items = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot "tts.json") | ConvertFrom-Json
foreach ($it in $items) {
  $p = Join-Path $PSScriptRoot ("c{0:d4}.wav" -f $it.i)
  $s.SetOutputToWaveFile($p)
  $s.Speak($it.t)
}
$s.SetOutputToNull(); $s.Dispose()
""".strip(), encoding="utf-8")
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                    "-File", str(ps)], check=True, capture_output=True)

    out, k = [], 0
    for g in cues_per_slide:
        grp = []
        for _ in g:
            k += 1
            grp.append(WORK / f"c{k:04d}.wav")
        out.append(grp)
    return out


def wav_sec(p):
    with wave.open(str(p)) as w:
        return w.getnframes() / w.getframerate()


# ── 4. 合成 H.265 MP4 ──────────────────────────────────────────────────
def build(frames, secs, wav_groups, srt):
    import imageio_ffmpeg
    ff = imageio_ffmpeg.get_ffmpeg_exe()

    lst = WORK / "concat.txt"
    lines = []
    for f, s in zip(frames, secs):
        lines.append("file '%s'" % f.name)
        lines.append("duration %.3f" % s)
    lines.append("file '%s'" % frames[-1].name)     # concat 規格要求重複最後一張
    lst.write_text("\n".join(lines), encoding="utf-8")

    # 燒進畫面的字幕。字型一定要指名微軟正黑體 ——
    # 不指名 libass 會挑到沒有中文字的字型，整排變成方框。
    # 投影片縮到 1920x960 擺上面，下面留 120px 色塊放字幕 ——
    # 字幕直接壓在投影片上會蓋住截圖與說明，SOP 反而看不清楚。
    style = ("FontName=Microsoft JhengHei,FontSize=26,"
             "PrimaryColour=&H00FFFFFF,OutlineColour=&H00000000,"
             "BorderStyle=1,Outline=1,Shadow=0,"
             "Alignment=2,MarginV=14")
    vf = ("scale=1920:960:flags=lanczos,"
          "pad=1920:1080:0:0:color=0x111318,"
          + "subtitles=%s:fontsdir=C\\\\:/Windows/Fonts"
          + ":force_style='%s'") % (srt.name, style)

    silent = WORK / "video.mp4"
    # ffmpeg 7：-r 和 -vsync vfr 互斥會直接報 Invalid argument，
    # 投影片影片本來就要固定 30fps，所以只給 -r。
    # cwd 設在 WORK，字幕路徑才不用處理 Windows 的磁碟機冒號。
    subprocess.run([ff, "-y", "-f", "concat", "-safe", "0", "-i", lst.name,
                    "-vf", vf, "-pix_fmt", "yuv420p",
                    "-c:v", "libx265", "-crf", "26", "-preset", "medium",
                    "-tag:v", "hvc1",      # 沒這個 tag，Windows 內建播放器打不開
                    "-r", "30", silent.name],
                   check=True, capture_output=True, cwd=WORK)

    wavs = [w for g in wav_groups for w in g]
    if not wavs:
        shutil.copy(silent, OUT_MP4)
        return

    alist = WORK / "audio.txt"
    alist.write_text("\n".join("file '%s'" % w.name for w in wavs),
                     encoding="utf-8")
    track = WORK / "audio.m4a"
    subprocess.run([ff, "-y", "-f", "concat", "-safe", "0", "-i", alist.name,
                    "-c:a", "aac", "-b:a", "128k", track.name],
                   check=True, capture_output=True, cwd=WORK)
    # 除了燒進畫面，另外掛一軌可開關的字幕
    subprocess.run([ff, "-y", "-i", silent.name, "-i", track.name,
                    "-i", srt.name,
                    "-map", "0:v", "-map", "1:a", "-map", "2:s",
                    "-c:v", "copy", "-c:a", "copy", "-c:s", "mov_text",
                    "-metadata:s:s:0", "language=cht", "-shortest",
                    str(OUT_MP4)], check=True, capture_output=True, cwd=WORK)


def main():
    cues = narration()
    n_cue = sum(len(g) for g in cues)
    print("投影片 %d 張，字幕 %d 句" % (len(cues), n_cue))
    frames = render_frames(len(cues))

    if NO_TTS:
        wav_groups = [[] for _ in cues]
        secs = [MIN_SEC + 1.2 * len(g) for g in cues]
    else:
        have = [WORK / ("c%04d.wav" % (i + 1)) for i in range(n_cue)]
        extra = WORK / ("c%04d.wav" % (n_cue + 1))
        if REUSE and all(w.exists() for w in have) and not extra.exists():
            print("  沿用既有旁白")
            wav_groups, k = [], 0
            for g in cues:
                wav_groups.append(have[k:k + len(g)]); k += len(g)
        else:
            print("  產生旁白（zh-TW，逐句）…")
            wav_groups = tts(cues)
        secs = [max(MIN_SEC, sum(wav_sec(w) for w in g) + PAD_SEC)
                for g in wav_groups]

    # 字幕時間軸：用每一句旁白的實際長度排，才對得準
    srt, t, idx = [], 0.0, 0
    for g, texts, dur in zip(wav_groups, cues, secs):
        at = t
        for j, txt in enumerate(texts):
            d = wav_sec(g[j]) if g else dur / max(1, len(texts))
            idx += 1
            srt.append("%d\n%s --> %s\n%s\n" % (
                idx, srt_time(at), srt_time(min(at + d, t + dur)), wrap(txt)))
            at += d
        t += dur
    srt_path = WORK / "sub.srt"
    srt_path.write_text("\n".join(srt), encoding="utf-8-sig")
    shutil.copy(srt_path, OUT_MP4.with_suffix(".srt"))

    print("  總長度 %.1f 分鐘，編碼中（字幕燒進畫面＋另掛一軌）…" % (sum(secs) / 60))
    build(frames, secs, wav_groups, srt_path)
    print("\n✓ %s  (%.1f MB)" % (OUT_MP4, OUT_MP4.stat().st_size / 1048576))
    print("✓ %s" % OUT_MP4.with_suffix(".srt"))


main()
