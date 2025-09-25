import os
import winreg
import vdf  # You might need to install this: pip install vdf


def get_steam_install_path():
    """Retrieves the Steam installation path from the Windows Registry."""
    try:
        # Try 64-bit registry path first
        key = winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Wow6432Node\\Valve\\Steam")
        install_path = winreg.QueryValueEx(key, "InstallPath")[0]
        winreg.CloseKey(key)
        return install_path
    except FileNotFoundError:
        try:
            # Fallback to 32-bit registry path
            key = winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Valve\\Steam")
            install_path = winreg.QueryValueEx(key, "InstallPath")[0]
            winreg.CloseKey(key)
            return install_path
        except FileNotFoundError:
            return None


def get_steam_library_paths():
    """Finds all Steam library paths, including the main install and additional folders."""
    library_paths = []
    steam_install_path = get_steam_install_path()

    if steam_install_path:
        # Look for additional library folders in libraryfolders.vdf
        libraryfolders_vdf_path = os.path.join(steam_install_path, "steamapps", "libraryfolders.vdf")
        if os.path.exists(libraryfolders_vdf_path):
            try:
                with open(libraryfolders_vdf_path, 'r', encoding='utf-8') as f:
                    library_data = vdf.load(f)

                # Iterate through the "LibraryFolders" section
                for key, value in library_data.get("libraryfolders", {}).items():
                    if key.isdigit() and "path" in value:
                        # Add the "steamapps" subdirectory for each additional library
                        library_paths.append(os.path.join(value["path"], "steamapps"))
            except Exception as e:
                print(f"Error parsing libraryfolders.vdf: {e}")

    return [os.path.normpath(path) for path in library_paths]

def get_steam_game_path(path_after_common):
    """Finds the full path to a game within Steam's library folders."""
    library_paths = get_steam_library_paths()
    for library_path in library_paths:
        game_path = os.path.join(library_path, 'common', path_after_common)
        if os.path.exists(game_path):
            return game_path
    return None