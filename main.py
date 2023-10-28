import os
import win32com.client


def delete_shortcut(path):
    """Delete an existing shortcut."""
    if os.path.exists(path):
        os.remove(path)


def create_shortcut(src_path, shortcut_path, new_target_appendix=""):
    """Create a shortcut and modify its Target path."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = src_path

    if new_target_appendix:
        shortcut.Arguments = new_target_appendix

    shortcut.Save()


def main():
    # Shortcut name to delete. This is the default name in an english version of Windows if a shortcut was previously created. Otherwise it will be 'League of Legends'.
    shortcut_name = "LeagueClient.exe - Shortcut.lnk"

    # Path to Desktop of the current user
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    shortcut_path = os.path.join(desktop_path, shortcut_name)

    existing_shortcut_path = shortcut_path

    delete_shortcut(existing_shortcut_path)

    new_shortcut_path = shortcut_path

    # Path of the application/exe for which we're creating the shortcut.
    src_path = r"C:\Riot Games\League of Legends\LeagueClient.exe"

    # Text to append to the Target property of the shortcut
    target_appendix = "--locale=en_US"

    create_shortcut(src_path, new_shortcut_path, target_appendix)


if __name__ == "__main__":
    main()
