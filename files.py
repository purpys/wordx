

def list_files_ext(path, ext):
    import os
    files=[path+os.path.sep+file for file in os.listdir(path) if file.endswith("."+ext) and not file.startswith("~") and not file.startswith(".")]
    return files
