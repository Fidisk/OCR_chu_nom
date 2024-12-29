import os

for i in range(1, 90):  # Loop from 1 to 89
    input_file = f"aaad.{i}.jpeg"
    command = f"python predict.py --input={input_file} [--print_image]"
    os.system(command)  # Execute the command
