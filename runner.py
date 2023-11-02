import scraping  # This will import your scraping.py file as a module
import subprocess

def run_exist():
    # Execute exist.py as a subprocess, as if it were run from the command line
    subprocess.run(['python', 'exist.py'], check=True)

if __name__ == "__main__":
    # Run the main function from scraping.py
    scraping.main()

    # After scraping.py finishes, run exist.py
    run_exist()