from colorama import init, Fore, Style
import os

# 初始化colorama，支持跨平台彩色输出
init(autoreset=True)


def print_aligned_banner():
    # 清屏（适配Windows/Linux/macOS）
    os.system('cls' if os.name == 'nt' else 'clear')

    # 定义PortScanner的ASCII图案行（需确保每行内容准确，此处以示例为准）
    banner_lines = [
        "____            _   ____",
        "|  _ \\ ___  _ __| |_/ ___|  ___ __ _ _ __  _ __   ___ _ __",
        "| |_) / _ \\| '__| __\\___ \\ / __/ _` | '_ \\| '_ \\ / _ \\ '__|",
        "|  __/ (_) | |  | |_ ___) | (_| (_| | | | | | | |  __/ |",
        "|_|   \\___/|_|   \\__|____/ \\___\\__,_|_| |_|_| |_|\\___|_|"
    ]

    # 计算所有行的最大长度（用于统一对齐）
    max_line_length = max(len(line) for line in banner_lines)

    # 打印彩色对齐的Banner
    print("\n" * 2)  # 顶部留白
    for i, line in enumerate(banner_lines):
        # 为每行分配不同颜色（循环使用青、绿、黄、品红、蓝）
        colors = [Fore.CYAN, Fore.GREEN, Fore.YELLOW, Fore.MAGENTA, Fore.BLUE]
        colored_line = colors[i % len(colors)] + line.ljust(max_line_length)  # 左对齐并补全长度
        print(colored_line)

    # 打印分隔线与作者信息
    divider = "-" * (max_line_length + 10)
    print("\n" + Fore.WHITE + divider)
    author_info = [
        f"{Fore.CYAN}Author: {Fore.WHITE}Bifish",
        f"{Fore.CYAN}GitHub: {Fore.WHITE}https://github.com/Bifishone"
    ]
    for info in author_info:
        print(info.center(max_line_length + 10))  # 作者信息居中对齐
    print(Fore.WHITE + divider + "\n" * 2)


if __name__ == "__main__":
    print_aligned_banner()
    input("按回车键退出...")  # 保持窗口，方便查看