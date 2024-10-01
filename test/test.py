import os

def get_dirs(path):
    dirs = [path]
    for item in os.listdir(path):
        full_path = f'{path}/{item}'
        if os.path.isdir(full_path) and '.docx' in [item[-5::] for item in os.listdir(full_path)]:
            dirs += get_dirs(full_path)
    
    return dirs

if __name__ == '__main__':
    target_dir = 'Docs'

    print(get_dirs(target_dir))
