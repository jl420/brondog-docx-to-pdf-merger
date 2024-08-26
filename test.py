import os

def get_dirs(path):
    dirs = [path]
    for item in os.listdir(path):
        full_path = f'{path}/{item}'
        if os.path.isdir(full_path):
            # print(f'{item} is a directory')
            dirs += get_dirs(full_path)
        else:
            # print(f'{item} is a file')
            pass
    
    return dirs

if __name__ == '__main__':
    target_dir = 'Docs'

    print(get_dirs(target_dir))
