import subprocess
import time
import os
import sys

def force_restart_explorer():
    """强制重启explorer.exe - 会有短暂黑屏效果"""
    print("正在强制重启explorer.exe...")
    print("注意：屏幕会短暂变黑，这是正常现象")
    
    try:
        # 方法1: 直接使用wmic强制终止explorer进程
        print("步骤1: 强制终止explorer进程...")
        subprocess.run([
            "wmic", "process", "where", "name='explorer.exe'", "delete"
        ], check=True, capture_output=True, text=True)
        
        # 等待进程完全结束
        print("步骤2: 等待进程完全结束...")
        time.sleep(3)
        
        # 重新启动explorer
        print("步骤3: 重新启动explorer...")
        subprocess.Popen("explorer.exe")
        
        # 再等待一下让系统稳定
        time.sleep(2)
        print("✓ explorer.exe 重启完成！")
        return True
        
    except subprocess.CalledProcessError:
        print("方法1失败，尝试备用方法...")
        return force_restart_explorer_backup()

def force_restart_explorer_backup():
    """备用方法：使用PowerShell强制重启"""
    try:
        print("使用PowerShell强制重启...")
        # 更激进的PowerShell命令
        cmd = """
        # 获取所有explorer进程并强制结束
        Get-Process -Name explorer -ErrorAction SilentlyContinue | ForEach-Object { 
            $_.Kill() 
        }
        
        # 等待3秒确保进程完全结束
        Start-Sleep -Seconds 3
        
        # 重新启动explorer
        Start-Process explorer.exe
        """
        
        subprocess.run([
            "powershell", "-ExecutionPolicy", "Bypass", "-Command", cmd
        ], check=True, capture_output=True, text=True)
        
        time.sleep(2)
        print("✓ 备用方法成功！")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"备用方法也失败了: {e}")
        return try_cmd_method()

def try_cmd_method():
    """最后的方法：使用CMD命令"""
    try:
        print("尝试CMD方法...")
        
        # 使用cmd的taskkill命令
        subprocess.run([
            "cmd", "/c", "taskkill /f /im explorer.exe & timeout /t 3 & start explorer.exe"
        ], check=True, capture_output=True, text=True)
        
        time.sleep(2)
        print("✓ CMD方法成功！")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"CMD方法失败: {e}")
        return False

def refresh_desktop():
    """刷新桌面 - 触发系统重新调整"""
    try:
        print("正在刷新桌面...")
        # 发送F5刷新桌面
        subprocess.run([
            "powershell", "-Command", 
            "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{F5}')"
        ], capture_output=True, text=True)
        
        # 也可以通过注册表方式强制刷新
        subprocess.run([
            "powershell", "-Command",
            "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters"
        ], capture_output=True, text=True)
        
        print("✓ 桌面刷新完成")
        
    except Exception as e:
        print(f"桌面刷新失败: {e}")

def main():
    print("=" * 50)
    print("桌面修复工具 - 专治远程桌面打乱桌面布局")
    print("=" * 50)
    print()
    
    # 警告用户
    print("⚠️  警告：")
    print("   - 此操作会导致屏幕短暂变黑（1-3秒）")
    print("   - 所有打开的文件夹窗口会被关闭")
    print("   - 请保存好当前工作后再继续")
    print()
    
    # 询问用户是否继续
    choice = input("确定要继续吗？(y/n): ").lower().strip()
    if choice != 'y':
        print("操作已取消")
        return
    
    print("\n开始修复桌面...")
    print("-" * 30)
    
    # 强制重启explorer
    if force_restart_explorer():
        print("\n等待系统稳定...")
        time.sleep(3)
        
        # 刷新桌面
        refresh_desktop()
        
        print("\n" + "=" * 50)
        print("✅ 桌面修复完成！")
        print("如果分辨率还没恢复，请等待几秒钟或手动调整显示设置")
        print("=" * 50)
    else:
        print("\n❌ 修复失败，请尝试以管理员权限运行此脚本")
        print("或者手动重启explorer.exe进程")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n用户取消操作")
    except Exception as e:
        print(f"\n发生错误: {e}")
    finally:
        input("\n按回车键退出...")
