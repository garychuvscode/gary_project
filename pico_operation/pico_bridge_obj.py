import os
import subprocess


class PicoBridge:
    def __init__(self):
        pass

    def make_dir(self, file_dir):
        """
        Create directories recursively on Pico.

        Args:
            file_dir (str): The directory path to be created.
        """
        try:
            dirs = file_dir.split("/")
            for d in dirs:
                subprocess.run(["ampy", "--port", "/dev/ttyACM0", "mkdir", d])
        except Exception as e:
            print(f"Error: {e}")

    def del_dir(self, file_dir):
        """
        Delete directory and its contents on Pico.

        Args:
            file_dir (str): The directory path to be deleted.
        """
        try:
            subprocess.run(["ampy", "--port", "/dev/ttyACM0", "rmdir", file_dir])
        except Exception as e:
            print(f"Error: {e}")

    def input_to_pico(self, file_dir):
        """
        Transfer files from computer to Pico.

        Args:
            file_dir (str): The directory path on the computer.
        """
        try:
            self.make_dir(file_dir)
            for root, dirs, files in os.walk(file_dir):
                for f in files:
                    file_path = os.path.join(root, f)
                    rel_path = os.path.relpath(file_path, file_dir)
                    pico_path = os.path.join("/", rel_path)
                    subprocess.run(
                        ["ampy", "--port", "/dev/ttyACM0", "put", file_path, pico_path]
                    )
        except Exception as e:
            print(f"Error: {e}")

    def output_from_pico(self, file_dir):
        """
        Transfer files from Pico to computer.

        Args:
            file_dir (str): The directory path on Pico.
        """
        try:
            subprocess.run(["ampy", "--port", "/dev/ttyACM0", "get", file_dir, "."])
        except Exception as e:
            print(f"Error: {e}")

    def dir_find(self, file_dir):
        """
        Find directory or file on Pico and transfer files to C:\pico_search.

        Args:
            file_dir (str): The directory or file to search for on Pico.
        """
        try:
            # Check if pico_search folder exists, if not, create it
            if not os.path.exists("C:/pico_search"):
                os.makedirs("C:/pico_search")

            # Search for file_dir on Pico and transfer files to C:\pico_search
            subprocess.run(
                ["ampy", "--port", "/dev/ttyACM0", "get", file_dir, "C:/pico_search"]
            )
        except Exception as e:
            print(f"Error: {e}")


# Testing Code
if __name__ == "__main__":
    # Instantiate PicoBridge
    pico_bridge = PicoBridge()

    # Test make_dir
    pico_bridge.make_dir("test_dir")

    # Test del_dir
    pico_bridge.del_dir("test_dir")

    # Test input_to_pico
    pico_bridge.input_to_pico("test_files")

    # Test output_from_pico
    pico_bridge.output_from_pico("/")

    # Test dir_find
    pico_bridge.dir_find("example_dir")
