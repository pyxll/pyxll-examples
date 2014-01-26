"""
Uninstall PyXLL from all Excel installations.

This script modifies the registry directly to remove references
to PyXLL from Excel.

Close all Excel sessions before running this script, as otherwise
Excel will re-write its settings when it closes so PyXLL will
still be installed.

PyXLL is removed from the following registry keys (both 32 and 64 bit):

- HKLM|HKCU/Software/Microsoft/Office/*/Excel/Options
- HKLM|HKCU/Software/Microsoft/Office/*/Excel/Add-in Manager
- HKLM|HKCU/Software/Microsoft/Office/*/Excel/Resiliency/DisabledItems

"""
import sys, os
import re
import logging
import _winreg as winreg

logging.basicConfig(level=logging.INFO)
_log = logging.getLogger(__name__)

_root_keys = {
    winreg.HKEY_CURRENT_USER : "HKEY_CURRENT_USER",
    winreg.HKEY_LOCAL_MACHINE : "HKEY_LOCAL_MACHINE",
}


def uninstall_all():
    """uninstalls PyXLL from all installed Excel versions"""
    for wow64_flags in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        for root in _root_keys.keys():
            try:
                flags = wow64_flags | winreg.KEY_READ
                office_root = winreg.OpenKey(root, r"Software\Microsoft\Office", 0, flags)
            except WindowsError:
                continue

            # look for all installed versions of Excel and uninstall PyXLL
            i = 0
            while True:
                try:
                    subkey = winreg.EnumKey(office_root, i)
                except WindowsError:
                    break

                match = re.match("^(\d+(?:\.\d+)?)$", subkey)
                if match:
                    office_version = match.group(1)
                    uninstall(office_root, office_version, wow64_flags)
                i += 1

            winreg.CloseKey(office_root)


def uninstall(office_root_key, office_version, wow64_flags):
    """Uninstalls PyXLL from a single Excel install"""
    # uninstall entries from \Software\Microsoft\Office\<version>\Excel\Options
    # (this is what Excel uses to determine what to load on start-up)
    options_key = None
    try:
        flags = wow64_flags | winreg.KEY_READ
        subkey = r"%s\Excel\Options" % office_version
        options_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
    except WindowsError:
        pass

    if options_key:
        _log.debug("Found %s Excel %s options keys" % (_get_arch(wow64_flags), office_version))
        pyxll_values = []
        try:
            i = 0
            while True:
                name, data, dtype = winreg.EnumValue(options_key, i)
                if "OPEN" in name and dtype == winreg.REG_SZ \
                and data.rstrip('"\'').lower().endswith("pyxll.xll"):
                    pyxll_values.append(name)
                i += 1
        except WindowsError:
            pass
        winreg.CloseKey(options_key)

        # if there were any pyxll keys found delete them
        if pyxll_values:
            _log.debug("Found PyXLL in %s Excel %s's options keys" % (_get_arch(wow64_flags), office_version))
            try:
                flags = wow64_flags | winreg.KEY_WRITE
                subkey = r"%s\Excel\Options" % office_version
                options_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
                for value in pyxll_values:
                    winreg.DeleteValue(options_key, value)
                winreg.CloseKey(options_key)
                _log.info("Deleted PyXLL from %s Excel %s's options" % (_get_arch(wow64_flags), office_version))
            except WindowsError:
                _log.error("Couldn't delete PyXLL keys from %s Excel %s's options; Write access not allowed." %
                           (_get_arch(wow64_flags), office_version))

    # uninstall entries from \Software\Microsoft\Office\<version>\Excel\Add-in Manager
    # (this is what Excel uses to list addins in the addin manager)
    addins_key = None
    try:
        flags = wow64_flags | winreg.KEY_READ
        subkey = r"%s\Excel\Add-in Manager" % office_version
        addins_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
    except WindowsError:
        pass

    if addins_key:
        _log.debug("Found %s Excel %s Addins" % (_get_arch(wow64_flags), office_version))
        pyxll_values = []
        try:
            i = 0
            while True:
                name, data, dtype = winreg.EnumValue(addins_key, i)
                filename = os.path.basename(name)
                if filename.lower() == "pyxll.xll":
                    pyxll_values.append(name)
                i += 1
        except WindowsError:
            pass
        winreg.CloseKey(addins_key)

        # if there were any pyxll keys found delete them
        if pyxll_values:
            _log.debug("Found PyXLL in %s Excel %s's Addins" % (_get_arch(wow64_flags), office_version))
            try:
                flags = wow64_flags | winreg.KEY_WRITE
                subkey = r"%s\Excel\Add-in Manager" % office_version
                addins_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
                for value in pyxll_values:
                    winreg.DeleteValue(addins_key, value)
                winreg.CloseKey(addins_key)
                _log.info("Deleted PyXLL from %s Excel %s's addins list" % (_get_arch(wow64_flags), office_version))
            except WindowsError:
                _log.error("Couldn't delete PyXLL keys from %s Excel %s's addins; Write access not allowed." %
                           (_get_arch(wow64_flags), office_version))

    # uninstall entries from \Software\Microsoft\Office\<version>\Excel\Resiliency\DisabledItems
    # (this is what Excel uses to list blacklist badly behaving addins)
    disabled_key = None
    try:
        flags = wow64_flags | winreg.KEY_READ
        subkey = r"%s\Excel\Resiliency\DisabledItems" % office_version
        disabled_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
    except WindowsError:
        pass

    if disabled_key:
        _log.debug("Found %s Excel %s disabled addins" % (_get_arch(wow64_flags), office_version))
        pyxll_values = []
        try:
            i = 0
            while True:
                name, data, dtype = winreg.EnumValue(disabled_key, i)
                if dtype == winreg.REG_BINARY:
                    value = data.decode("utf-16", "ignore")
                    if "pyxll.xll" in value:
                        pyxll_values.append(name)
                i += 1
        except WindowsError:
            pass
        winreg.CloseKey(disabled_key)

        # if there were any pyxll keys found delete them
        if pyxll_values:
            _log.debug("Found PyXLL in %s Excel %s's disabled addins" % (_get_arch(wow64_flags), office_version))
            try:
                flags = wow64_flags | winreg.KEY_WRITE
                subkey = r"%s\Excel\Resiliency\DisabledItems" % office_version
                disabled_key = winreg.OpenKey(office_root_key, subkey, 0, flags)
                for value in pyxll_values:
                    winreg.DeleteValue(disabled_key, value)
                winreg.CloseKey(addins_key)
                _log.info("Deleted PyXLL from %s Excel %s's disabled addins" % (_get_arch(wow64_flags), office_version))
            except WindowsError:
                _log.error("Couldn't delete PyXLL keys from %s Excel %s's disabled addins; Write access not allowed." %
                           (_get_arch(wow64_flags), office_version))


def _get_arch(flags):
    if flags & winreg.KEY_WOW64_64KEY:
        return "64 bit"
    elif flags & winreg.KEY_WOW64_32KEY:
        return "32 bit"
    return "unknown"


def main():
    uninstall_all()

if __name__ == "__main__":
    sys.exit(main())
