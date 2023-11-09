import scraping
import subprocess
# Create the subprocess
def run_charts():
    subprocess.run(['python', 'charts.py'], check=True)

if __name__ == "__main__":
    # Run the main function from scraping.py
    scraping.setup_gui()  # This calls the run_scraping function from the scraping module

    # After scraping.py finishes, this will run the function that creates charts within xlsx file
    run_charts()
