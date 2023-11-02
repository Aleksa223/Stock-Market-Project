import scraping 
import subprocess

def run_charts():
    # Execute exist.py as a subprocess, as if it were run from the command line
    subprocess.run(['python', 'charts.py'], check=True)

if __name__ == "__main__":
    # Run the main function from scraping.py
    scraping.main()

    # After scraping.py finishes, run exist.py
    run_charts()
