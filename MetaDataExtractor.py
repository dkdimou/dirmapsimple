import datetime
import win32security

class MetaDataExtractor:
    def __init__(self):
        self.sid_cache = {}

    SID_TYPE_DICT = {
        win32security.SidTypeUser: "User",
        win32security.SidTypeGroup: "Group",
        win32security.SidTypeDomain: "Domain",
        win32security.SidTypeAlias: "Alias",
        win32security.SidTypeWellKnownGroup: "WellKnownGroup",
        win32security.SidTypeDeletedAccount: "DeletedAccount",
        win32security.SidTypeInvalid: "Invalid",
        win32security.SidTypeUnknown: "Unknown"
    }

    def get_file_security_info(self, path):
        sd = win32security.GetFileSecurity(str(path), win32security.OWNER_SECURITY_INFORMATION)
        owner_sid = sd.GetSecurityDescriptorOwner()
        # Convert owner_sid to a string so it can be used as a dictionary key
        owner_sid_str = str(owner_sid)
        if owner_sid_str not in self.sid_cache:
            name, domain, type = win32security.LookupAccountSid(None, owner_sid)
            self.sid_cache[owner_sid_str] = (name, domain, type)
        else:
            name, domain, type = self.sid_cache[owner_sid_str]

        stat = path.stat()

        def format_size(size):
            # size is in bytes
            for unit in ['bytes', 'KB', 'MB', 'GB', 'TB']:
                if size < 1024.0:
                    return f"{size:3.1f} {unit}"
                size /= 1024.0

        metadata_ = {
            'Owner': name,
            'Domain': domain,
            'Sid Type': self.SID_TYPE_DICT.get(type, "Unknown"),
            "Created": datetime.datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
            "Accessed": datetime.datetime.fromtimestamp(stat.st_atime).strftime('%Y-%m-%d %H:%M:%S'),
            "Modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            "Size": format_size(stat.st_size)
        }
        return metadata_
