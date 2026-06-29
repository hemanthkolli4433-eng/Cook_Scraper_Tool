COOK COUNTY GITHUB AUTOMATION - SELF-HOSTED RUNNER ONLY

Why this version:
GitHub cloud/headless browser is failing after 3 attempts because the Cook County site is not returning the result page from GitHub cloud. This folder uses your own Windows computer as the GitHub runner, so the browser runs from your internet and visible Chrome.

FILES TO UPLOAD TO GITHUB:
Cook_Search_Tool.py
Cook_input.txt
requirements.txt
.gitattributes
run_local_visible.bat
run_local_headless_test.bat
.github/workflows/cook_daily_run_self_hosted.yml
README_SELF_HOSTED_STEPS.txt

STEP 1 - Test locally first:
1. Double-click run_local_visible.bat
2. Chrome should open visibly.
3. If data saves to Cook_output.xlsx locally, continue.

STEP 2 - Push these files to GitHub:
git add .
git commit -m "Use self hosted Windows runner for Cook County scraper"
git push

STEP 3 - Add self-hosted runner in GitHub:
1. Open your repo: https://github.com/hemanthkolli4433-eng/Cook_Scraper_Tool
2. Go to Settings
3. Left side: Actions
4. Click Runners
5. Click New self-hosted runner
6. Select Windows and x64
7. GitHub will show exact commands. Copy and run those commands in PowerShell.

IMPORTANT:
When configuring the runner, DO NOT install it as a service if you need visible Chrome.
Run it from PowerShell using:
.\run.cmd
Leave that PowerShell window open. It should say:
Listening for Jobs

STEP 4 - Run the workflow:
1. GitHub repo -> Actions
2. Select "Cook County Daily Run - Self Hosted Windows"
3. Click Run workflow
4. Wait until it finishes
5. Download artifact named Cook_output_self_hosted_<run number>

STEP 5 - Automatic schedule:
The workflow is scheduled daily at 10:30 AM India time.
Your computer must be ON, internet connected, and the runner PowerShell window must be running.

Do not run the cloud headless workflow for this website. Use only the self-hosted Windows workflow.
