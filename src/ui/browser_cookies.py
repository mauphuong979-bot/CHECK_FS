"""Script phía trình duyệt cho cookie "Ghi nhớ đăng nhập".

Streamlit chỉ đọc được cookie (``st.context.cookies``) chứ không ghi được, nên cookie được đặt/xóa bằng
script chạy trong iframe cùng origin. Một số máy chủ (ví dụ Streamlit Community Cloud) không chuyển cookie tới
ứng dụng; khi đó "cầu nối" đọc cookie trong trình duyệt và gõ token vào một ô nhập ẩn để ứng dụng nhận được.
Token không bao giờ được đưa lên URL.
"""

from __future__ import annotations

import json

COOKIE_NAME = "checkfs_remember"
BRIDGE_INPUT_KEY = "checkfs_remember_bridge"
NO_TOKEN = "-"


def set_cookie_script(token: str, max_age_seconds: int) -> str:
    return _cookie_script(f"{COOKIE_NAME}={token}; max-age={max_age_seconds}")


def clear_cookie_script() -> str:
    return _cookie_script(f"{COOKIE_NAME}=; max-age=0")


def _cookie_script(cookie: str) -> str:
    return f"""<script>
        var secure = window.parent.location.protocol === 'https:' ? '; Secure' : '';
        window.parent.document.cookie = {json.dumps(cookie)} + '; path=/; SameSite=Strict' + secure;
    </script>"""


def bridge_script() -> str:
    """Đọc cookie trong trình duyệt và gõ token (hoặc ``NO_TOKEN``) vào ô nhập ẩn ``BRIDGE_INPUT_KEY``."""
    return f"""<script>
    (function () {{
        var page = window.parent;
        try {{
            var match = page.document.cookie.match(/(?:^|; ){COOKIE_NAME}=([^;]+)/);
            var token = match ? match[1] : {json.dumps(NO_TOKEN)}, tries = 0, sends = 0;
            function type(input) {{
                var setValue = Object.getOwnPropertyDescriptor(page.HTMLInputElement.prototype, "value").set;
                setValue.call(input, token);
                input.dispatchEvent(new page.Event("input", {{ bubbles: true }}));
                ["keydown", "keypress", "keyup"].forEach(function (kind) {{
                    input.dispatchEvent(new page.KeyboardEvent(kind, {{ key: "Enter", code: "Enter",
                        keyCode: 13, which: 13, charCode: 13, bubbles: true }}));
                }});
            }}
            (function send() {{
                var input = page.document.querySelector(".st-key-{BRIDGE_INPUT_KEY} input");
                if (input) {{
                    setTimeout(function () {{
                        type(input);
                        if (++sends < 6) setTimeout(function () {{
                            var box = input.closest("[data-stale]");
                            if (input.isConnected && !(box && box.getAttribute("data-stale") === "true")) send();
                        }}, 700);
                    }}, 150);
                }} else if (++tries < 30) {{
                    setTimeout(send, 100);
                }}
            }})();
        }} catch (error) {{
            console.log("remember-me bridge:", error);
        }}
    }})();
    </script>"""
