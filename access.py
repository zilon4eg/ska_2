from winsys import fs
import os
from tqdm import tqdm


def directory_access(directory_path):
    directory = fs.dir(directory_path)
    all_access = []
    for ace in directory.security().dacl:
        user = str(ace.trustee)[str(ace.trustee).find('\\') + 1:]
        access_flags = fs.FILE_ACCESS.names_from_value(ace.access)
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
            access.append(user)
            access.append(directory_path[directory_path.rfind('\\') + 1:])
            access.reverse()
            for i in range(5):
                if len(access) < 7:
                    access.append(None)
            all_access.append(access)
    return all_access


def dir_list(root_path):
    return list(item for item in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, item)))


def directories_access(root_path):
    access = []
    for directory in tqdm(dir_list(root_path), 'Анализируем права доступа: '):
        path = f'{root_path}\\{directory}'
        for dir in directory_access(path):
            access.append(dir)
    return access

