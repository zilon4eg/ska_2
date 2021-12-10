from winsys import fs
import winsys
import os
from pprint import pprint


def directory_access(dir_path):
    directory = winsys.fs.dir(dir_path)
    all_access = []
    for ace in directory.security().dacl:
        user = str(ace.trustee)[str(ace.trustee).find('\\') + 1:]
        access_flags = winsys.fs.FILE_ACCESS.names_from_value(ace.access)
        if user not in ['Администратор', 'Прошедшие проверку', 'СИСТЕМА', 'Администраторы', 'Пользователи', 'Администраторы домена']:

            access = []
            # ===(чтение)===
            read = ['ALL_ACCESS', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'READ_CONTROL']
            # ===(запись)===
            write = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'SYNCHRONIZE']
            # ===(чтение и выполнение + список содержимого папки + чтение)===
            execution = ['ALL_ACCESS', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'READ_CONTROL']
            # ===(изменение + чтение и выполнение + список содержимого папки + чтение + запись)===
            modification = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'DELETE', 'READ_CONTROL']
            # ===(полный доступ)===
            full_access = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'DELETE_CHILD', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'DELETE', 'READ_CONTROL', 'WRITE_DAC', 'WRITE_OWNER']

            # set.issubset(other) или set <= other - все элементы set принадлежат other
            if set(full_access).issubset(set(access_flags)):
                access.append('Полный доступ')
            if set(modification).issubset(set(access_flags)):
                access.append('Изменение')
            if set(write).issubset(set(access_flags)):
                access.append('Запись')
            if set(execution).issubset(set(access_flags)):
                access.append('Чтение и выполнение')
            if set(read).issubset(set(access_flags)):
                access.append('Чтение')
            access.reverse()
            all_access.append({user: access})
    return all_access


def dir_list(root_path):
    return [item for item in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, item))]


def directories_access(root_path):
    access = {}
    for directory in dir_list(root_path):
        path = f'\\\\fs\SHARE\Documents\{directory}'
        access[directory] = directory_access(path)
    return access


if __name__ == '__main__':
    root = r'\\fs\SHARE\Documents'
    pprint(directories_access(root))



