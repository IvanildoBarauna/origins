from oletools.olevba import VBA_Parser
import os

files = os.listdir("/Users/ivbarauna/repos/origins/source_files")

def extract_full_code():
    for file in files: 
        vba_parser = VBA_Parser("/Users/ivbarauna/repos/origins/source_files/" + file)
        if vba_parser.detect_vba_macros():
            for (
                filename,
                stream_path,
                vba_filename,
                vba_code,
            ) in vba_parser.extract_all_macros():
                print(f"VBA Code Founded: {vba_filename}")
                with open(f"new_files/{vba_filename}", "w") as f:
                    f.write(vba_code)
        else:
            print("Nenhum código VBA encontrado.")
            
extract_full_code()        