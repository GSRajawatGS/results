# result_jobs.py - the only file you need to edit for result_scraper.py.
# Keep it in the same folder as result_scraper.py. Lines starting with # are comments.

# ---- The result site -----------------------------------------------------
START_URL = "https://univindia.com/JwgResult/adselcolclassnew.php"
EXAM_YEAR = "2023"       # exactly as the "Years" dropdown lists it
SESSION = "JUN-2023"     # used for the Excel file names, and to pick the right exam if a code appears twice

# ---- Folders --------------------------------------------------------------
# PDFs saved by mponline.py (attestation sheets of the session)
ATT_FOLDER = r"C:\Users\acer\Documents\JUN_2023\JUN_2023_ATTESTATION\G107-LAWCOLLEGE,MORENA"
# Excel files are written here
OUT_FOLDER = r"C:\Users\acer\Documents\JUN_2023\JUN_2023_RESULTS"

# ---- Which classes ----------------------------------------------------------
# Write them the way you would say them, or type the exam code itself (ALL323, LL335 ...).
# Leave the list empty to be asked each time the script runs.
CLASSES = [
    # "llb 3 sem",
    # "llb 3 sem atkt",
    # "llb 5 sem",
]

# course name -> (letters in the EXCODE, course-duration digit)   e.g. LL + 3 (3YDC) + year + semester
COURSE_MAP = {
    "llb": ("LL", "3"),
}

# Which attestation files belong to the main exam and which to the ATKT exam
# (the Status part of the PDF file name, e.g. 3SEM_REGULAR_EXR_LLBC254.pdf)
MAIN_STATUSES = ["REGULAR", "EXR"]
ATKT_STATUSES = ["ATKT"]

# ---- Options ------------------------------------------------------------------
HEADLESS = False         # True = hide the browser window
DELAY = (0.8, 1.6)       # pause (seconds) between two roll numbers, so the site is not hammered

# ---- Automatic mode (python result_scraper.py auto / watch) ---------------------
# It looks in this folder for session folders like JUN_2026\JUN_2026_ATTESTATION (made by mponline.py),
# works out the session, exam year and every class (main + ATKT) from the PDF names, and writes
#   <DOCS_ROOT>\JUN_2026\JUN_2026_RESULTS\JUN_2026_<exam code>.xlsx
DOCS_ROOT = None                 # None = your My Documents folder
WATCH_MINUTES = 30               # 'watch' mode: how often to look for new PDFs
RETRY_NOT_FOUND_HOURS = 24       # roll numbers the site did not find are asked again after this long
RETRY_MISSING_CODE_HOURS = 6     # an exam that is not on the result list yet is looked for again after this long
