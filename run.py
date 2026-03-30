"""Startup script for Excel Assistant web application."""

import uvicorn


def main():
    """Run the Excel Assistant web server."""
    print("=" * 50)
    print("  📊 Excel助手 - 智能Excel处理平台")
    print("=" * 50)
    print()
    print("  启动 Web 服务器...")
    print("  访问地址: http://localhost:8000")
    print()
    print("  提示: 请确保已配置 .env 文件中的 LLM API 密钥")
    print("  按 Ctrl+C 停止服务器")
    print("=" * 50)
    print()

    uvicorn.run(
        "backend.app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
    )


if __name__ == "__main__":
    main()
