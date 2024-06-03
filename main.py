from oletools.olevba import VBA_Parser
import os
import datetime

files = os.listdir("source_files/")

def extract_full_code():
    for file in files:
        vba_parser = VBA_Parser("source_files/" + file)
        if vba_parser.detect_vba_macros():
            for (
                filename,
                stream_path,
                vba_filename,
                vba_code,
            ) in vba_parser.extract_all_macros():
                print(f"VBA Code Founded: {vba_filename}")
                os.makedirs(f"new_files/{file.split('.')[0]}", exist_ok=True)
                
                with open(f"new_files/{file.split('.')[0]}/{vba_filename}", "w") as f:
                    f.write(vba_code)
        else:
            print("Nenhum código VBA encontrado.")
            
extract_full_code()        