#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OneDrive 直链解析工具 - 网页版

支持：
- 个人版（onedrive.live.com / 1drv.ms）
- 海外企业版（*.sharepoint.com / *.my.sharepoint.com）
- 世纪互联版（*.sharepoint.cn / *.my.sharepoint.cn）

优化的世纪互联和企业版解析：
- 直接生成 /_layouts/52/download.aspx?share= 格式链接
- 避免需要登录的重定向链接

使用方法：
python OneDrive.py
"""

from __future__ import annotations

import re
import json
import os
import webbrowser
import threading
import time
from datetime import datetime
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse

import requests
from flask import Flask, render_template, request, jsonify


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/126.0 Safari/537.36"
)


class OneDriveLinkError(Exception):
    pass


def normalize_url(original_url: str, timeout: float) -> str:
    """展开 1drv.ms 等短链，返回实际的分享页 URL。"""
    try:
        resp = requests.head(
            original_url,
            allow_redirects=True,
            timeout=timeout,
            headers={"User-Agent": USER_AGENT},
        )
        if resp.status_code in (405, 403) or resp.is_redirect:
            resp = requests.get(
                original_url,
                allow_redirects=True,
                timeout=timeout,
                headers={"User-Agent": USER_AGENT},
                stream=True,
            )
        return resp.url
    except requests.RequestException as exc:
        raise OneDriveLinkError(f"无法展开链接: {exc}") from exc


def is_onedrive_supported_host(url: str) -> bool:
    host = urlparse(url).hostname or ""
    host = host.lower()
    if not host:
        return False
    return any(
        host.endswith(suffix)
        for suffix in (
            "1drv.ms",
            "onedrive.live.com",
            "sharepoint.com",
            "my.sharepoint.com",
            "sharepoint.cn",
            "my.sharepoint.cn",
        )
    )


def is_folder_link(url: str) -> bool:
    """检测是否为文件夹分享链接。"""
    parsed = urlparse(url)
    path_lower = (parsed.path or "").lower()
    
    # SharePoint 路径包含 ":f:" 基本可判定为文件夹
    if ":f:" in path_lower:
        return True
    
    # 检查是否为文档库根目录等文件夹类型
    if ":b:" in path_lower or "forms/allitems.aspx" in path_lower:
        return True
    
    return False


def convert_sharepoint_to_direct_link(url: str) -> str:
    """
    将 SharePoint 分享链接转换为直接下载链接
    
    按照用户提供的转换规则:
    输入: https://xxxx-my.sharepoint.cn/:u:/g/personal/xxxx_xxxx_partner_onmschina_cn/XY16c5xGfSxDn91zeEE3A2UBVpoPOzfNzUEDW6dQ
    输出: https://xxxx-my.sharepoint.cn/personal/xxxx_xxxx_partner_onmschina_cn/_layouts/52/download.aspx?share=XY16c5xGfSxDn91zeEE3A2UBVpoPOzfNzUEDW6dQ
    """
    parsed = urlparse(url)
    domain = parsed.netloc
    path = parsed.path
    
    # 只处理标准的 SharePoint 分享链接格式
    # /:u:/g/personal/用户名/分享token
    match = re.search(r'/:u:/g/personal/([^/]+)/([^/?#]+)', path)
    if match:
        personal_user = match.group(1)
        share_token = match.group(2)
        # 直接按照用户提供的格式转换
        direct_url = f"https://{domain}/personal/{personal_user}/_layouts/52/download.aspx?share={share_token}"
        return direct_url
    
    # 如果不是标准格式，提示用户使用正确的分享链接
    raise OneDriveLinkError(
        f"请使用标准的 OneDrive/SharePoint 分享链接。\n"
        f"支持格式: https://domain/:u:/g/personal/用户名/分享token\n"
        f"当前链接: {url}"
    )


def convert_personal_onedrive_to_direct_link(url: str) -> str:
    """
    将个人版 OneDrive 链接转换为真正的无登录直链
    
    使用 dlink.host 服务，通过提取 redeem 参数生成直链
    这是目前最有效的个人版OneDrive直链生成方法
    """
    parsed = urlparse(url)
    query_params = dict(parse_qsl(parsed.query))
    
    # 方法1: 提取 redeem 参数，使用 dlink.host 服务
    if 'redeem' in query_params:
        redeem_value = query_params['redeem']
        dlink_url = f"https://dlink.host/1drv/{redeem_value}"
        return dlink_url
    
    # 方法2: 如果有 resid 和 authkey，尝试传统格式
    if 'resid' in query_params and 'authkey' in query_params:
        file_id = query_params['resid']
        authkey = query_params['authkey']
        return f"https://onedrive.live.com/download?resid={file_id}&authkey={authkey}"
    
    # 回退到传统方法
    return ensure_download_param_fallback(url)


def ensure_download_param_fallback(url: str) -> str:
    """传统的 download=1 方法（备用）"""
    parsed = urlparse(url)
    query_pairs = dict(parse_qsl(parsed.query))
    
    if "download" not in {k.lower() for k in query_pairs.keys()}:
        query_pairs["download"] = "1"
    
    new_query = urlencode(query_pairs, doseq=True)
    new_url = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            parsed.fragment,
        )
    )
    return new_url


def parse_onedrive_direct_link(input_url: str, timeout: float = 15.0) -> str:
    if not re.match(r"^https?://", input_url, flags=re.IGNORECASE):
        raise OneDriveLinkError("请输入以 http:// 或 https:// 开头的分享链接")

    # 先检查原始链接是否为标准格式
    if not is_onedrive_supported_host(input_url):
        raise OneDriveLinkError("不支持的链接域名：仅支持 OneDrive/SharePoint/世纪互联分享链接")

    parsed = urlparse(input_url)
    host = parsed.netloc.lower()
    
    # 判断是否为 SharePoint（企业版或世纪互联）
    if any(domain in host for domain in ['sharepoint.com', 'sharepoint.cn', 'my.sharepoint']):
        # 对于 SharePoint，直接解析原始链接，不要跟随重定向
        return convert_sharepoint_to_direct_link(input_url)
    else:
        # 对于个人版 OneDrive，先展开短链再处理
        expanded = normalize_url(input_url, timeout=timeout)
        
        # 对于个人版，使用传统方法并添加说明
        direct_link = convert_personal_onedrive_to_direct_link(expanded)
        
        # 添加提示信息
        if 'dlink.host/1drv/' in direct_link:
            print(f"🎉 使用 dlink.host 服务生成真正的无登录直链！")
            print(f"💡 通过提取 redeem 参数实现，这是最有效的方法")
        elif 'onedrive.live.com/download?resid=' in direct_link:
            print(f"✨ 生成了优化的直链格式")
        else:
            print(f"⚠️  回退到传统格式，可能需要登录")
        
        return direct_link


# ============= Flask 网页应用 =============

app = Flask(__name__)

# 历史记录文件路径
HISTORY_FILE = "onedrive_history.json"

def load_history():
    """加载历史记录"""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

def save_history_to_file(history_data):
    """保存历史记录到文件"""
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存历史记录失败: {e}")
        return False

@app.get("/")
def index_get():
    return render_template("index.html", result=None, error=None, form={})

@app.post("/")
def index_post():
    share_url: str = request.form.get("share_url", "").strip()

    if not share_url:
        return render_template(
            "index.html",
            result=None,
            error="请填写分享链接",
            form={"share_url": share_url},
        )

    try:
        direct = parse_onedrive_direct_link(share_url, timeout=15.0)
        return render_template(
            "index.html",
            result=direct,
            error=None,
            form={"share_url": share_url},
        )
    except OneDriveLinkError as exc:
        return render_template(
            "index.html",
            result=None,
            error=str(exc),
            form={"share_url": share_url},
        )

@app.post("/save_history")
def save_history():
    """保存历史记录的API端点"""
    try:
        data = request.get_json()
        if not data or 'url' not in data:
            return jsonify({'success': False, 'error': '缺少必要参数'})
        
        # 加载现有历史记录
        history = load_history()
        
        # 创建新的历史记录项
        history_item = {
            'id': data.get('id', int(datetime.now().timestamp() * 1000)),
            'url': data['url'],
            'remark': data.get('remark', '未命名文件'),
            'timestamp': data.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            'original_url': data.get('original_url', ''),  # 可选：保存原始链接
        }
        
        # 避免重复
        existing_index = next((i for i, item in enumerate(history) if item['url'] == history_item['url']), -1)
        if existing_index != -1:
            history[existing_index] = history_item  # 更新现有记录
        else:
            history.insert(0, history_item)  # 添加到开头
        
        # 限制记录数量
        if len(history) > 100:
            history = history[:100]
        
        # 保存到文件
        if save_history_to_file(history):
            return jsonify({'success': True, 'message': '历史记录已保存'})
        else:
            return jsonify({'success': False, 'error': '保存失败'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.get("/history")
def get_history():
    """获取历史记录的API端点"""
    try:
        history = load_history()
        return jsonify({'success': True, 'data': history})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


def open_browser():
    """延迟打开浏览器"""
    time.sleep(0.5)  # 等待Flask服务启动
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == "__main__":
    print("启动 OneDrive 直链解析工具: http://127.0.0.1:5000")
    
    # 只在主进程中打开浏览器，避免调试模式重启时重复打开
    if os.environ.get('WERKZEUG_RUN_MAIN') != 'true':
        print("正在自动打开浏览器...")
        # 在后台线程中打开浏览器
        threading.Thread(target=open_browser, daemon=True).start()
    
    # 启动Flask应用
    app.run(host="127.0.0.1", port=5000, debug=True)