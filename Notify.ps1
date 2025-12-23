# ============================================================
# Toast Notification Helper
# ============================================================

param(
    [string]$Title = "System Notification",
    [string]$Message = "Action completed."
)

# Load Windows Runtime
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]

$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

$toastTextElements = $template.GetElementsByTagName("text")
$toastTextElements.Item(0).AppendChild($template.CreateTextNode($Title)) | Out-Null
$toastTextElements.Item(1).AppendChild($template.CreateTextNode($Message)) | Out-Null

$toast = [Windows.UI.Notifications.ToastNotification]::new($template)
$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("WPAS Automation")
$notifier.Show($toast)