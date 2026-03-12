#!/usr/bin/env python3
"""
桌面文件计数器
统计桌面目录下的所有文件数量
"""

import os
import sys
from pathlib import Path
from collections import Counter


def get_desktop_path():
    """获取桌面路径"""
    if len(sys.argv) > 1:
        return sys.argv[1]

    # 尝试从环境变量获取
    if 'USERPROFILE' in os.environ:
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        if os.path.exists(desktop):
            return desktop

    # 尝试从 HOME 获取
    if 'HOME' in os.environ:
        desktop = os.path.join(os.environ['HOME'], 'Desktop')
        if os.path.exists(desktop):
            return desktop

    # 尝试直接访问当前用户的桌面
    username = os.getlogin() if hasattr(os, 'getlogin') else 'unknown'
    desktop = f'C:\\Users\\{username}\\Desktop'
    if os.path.exists(desktop):
        return desktop

    return None


def count_files(directory):
    """
    统计目录下的文件数量

    Args:
        directory: 要统计的目录路径

    Returns:
        dict: 包含文件统计信息的字典
    """
    if not directory or not os.path.exists(directory):
        return {
            'error': f'目录不存在: {directory}'
        }

    file_count = 0
    dir_count = 0
    extensions = Counter()
    largest_file = {'name': '', 'size': 0}
    total_size = 0

    try:
        for root, dirs, files in os.walk(directory):
            # 统计文件
            for file in files:
                file_count += 1
                file_path = os.path.join(root, file)
                file_ext = Path(file).suffix.lower()
                extensions[file_ext] += 1

                # 统计文件大小
                try:
                    file_size = os.path.getsize(file_path)
                    total_size += file_size

                    # 记录最大的文件
                    if file_size > largest_file['size']:
                        largest_file = {
                            'name': file_path,
                            'size': file_size
                        }
                except (OSError, PermissionError):
                    pass

            # 统计目录（不包括根目录）
            if root != directory:
                dir_count += len(dirs)
            else:
                dir_count += len(dirs)

    except PermissionError as e:
        return {
            'error': f'权限不足: {e}'
        }
    except Exception as e:
        return {
            'error': f'发生错误: {e}'
        }

    # 格式化文件大小
    def format_size(size_bytes):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f'{size_bytes:.2f} {unit}'
            size_bytes /= 1024.0
        return f'{size_bytes:.2f} TB'

    return {
        'directory': directory,
        'total_files': file_count,
        'total_directories': dir_count,
        'extensions': dict(extensions),
        'largest_file': {
            'name': os.path.basename(largest_file['name']) if largest_file['name'] else 'N/A',
            'path': largest_file['name'],
            'size': format_size(largest_file['size'])
        },
        'total_size': format_size(total_size),
        'success': True
    }


def main():
    """主函数"""
    desktop_path = get_desktop_path()

    if not desktop_path:
        print("错误：无法确定桌面路径，请手动指定桌面路径")
        print(f"用法: python {sys.argv[0]} [桌面路径]")
        sys.exit(1)

    print(f"正在统计桌面文件: {desktop_path}")
    print("-" * 60)

    result = count_files(desktop_path)

    if 'error' in result:
        print(f"错误: {result['error']}")
        sys.exit(1)

    # 打印结果
    print("\n📊 桌面文件统计结果\n")
    print(f"📁 统计目录: {result['directory']}")
    print(f"📄 文件总数: {result['total_files']}")
    print(f"📂 目录总数: {result['total_directories']}")
    print(f"💾 总大小: {result['total_size']}")

    if result['extensions']:
        print("\n📋 文件类型分布:")
        # 按数量排序，取前10个
        sorted_extensions = sorted(
            result['extensions'].items(),
            key=lambda x: x[1],
            reverse=True
        )[:10]

        for ext, count in sorted_extensions:
            ext_name = ext if ext else '(无扩展名)'
            percentage = (count / result['total_files'] * 100)
            bar_length = int(percentage / 2)
            bar = '█' * bar_length
            print(f"  {ext_name:20s} {count:5d} ({percentage:5.1f}%) {bar}")

    if result['largest_file']['name'] != 'N/A':
        print(f"\n📏 最大文件: {result['largest_file']['name']} ({result['largest_file']['size']})")

    print("-" * 60)
    print("\n✅ 统计完成！")


if __name__ == '__main__':
    main()
