How does TEA work?

TEA - Technological Evidence Aquisition was originally started in 2009 in response to the Microsoft COFEE shortcomings at that time.
The concerns were:
COFEE was closed source, and therefor you could not testify to what was being done
COFEE was extremely difficult to get ahold of for Law Enforcement, and therefore not used or updated often
Versions at that time could not diferentiate between windows 32 bit and 64 bit, or windows XP and VISTA, therefore it failed a lot.

TEA originally was a batch file which would prompt you each step of the way.
Eventually I wanted to automate the process using something like a USB Rubberducky or similar Human Machine Interface (HMI) emulator.
In experimenting with this, I realized a CONFIGURATION file could be made ahead of time and then the process would be automated, hence BREW.


I worked on the tool from 2009-2014, and it was successfully used on approximately 20 search scenes. None of the evidence was ever called into question into court.

In 2021 I was asked to take over the Computer Forensic course at the Canadian Police College. I had expected other tools to have come out to render TEA obsolete,
however to my disappointment there are still not many first responder tools that take into account all the forensic artefacts we care about in Law Enforcement.

The Code is a bit of a frankenstein, with some parts existing but never being activated. The code could be optimized.

While the code could be made into Powershell or an Executable, it remains a batch file for the following reasons:
- Some computers block powershell. As you never know what your target is, batch scripts are less likely to be blocked.
- Older OS's dont have powershell, yet most these tools will work on 32bit XP
- An exececutable is not "open" for easy review and therefore more difficult to prove what happened each step of the way incase of an error.


Key features:
- Trusted Executables are run locally. So instead of running the local "net" command, you'd run "!!net.exe"  
- Windows Vista started blocking certian apps from running off a USB (like !!net.exe) so the local version was run. Initially I wasd keeping up to date with all HASHES of known good tools 
see - !hashlist.tea, but this became an onerous tasks, and I have yet to encounter an infected local tool that still functioned, so for now this is disabled. If this tool becomes popular and not protected under TLP,
people could hijack the tools on their system to prevent this tool from working, at that point the hashlist should be updated and re-enabled. 
Then we need to add logic that if a hash-mismatch is detected, try run a local alternative off the USB or just skip the tool.
- Most Forensic tools (FTKImager, looking at you) do not copy out the registry .LOG1 and .LOG2 files, resulting in only backing up DITRY hives, this usually results in incorrect and old date/time stamps. Getting the log files allow for clean hives to be rebuilts with more accurate times.
- Automatically will run as Administrator in the correct folder. Right clicking-run as administrator on the older tea's (or any batch) will automatically run in Windows\system32 and prevent the local tools from running. 
trying to determine the drive letter assigned to the USB with TEA on it was challenging unless you actually named the USB a specific name. This new code makes it simpler, just copy the TEA folder to any USB and run it.
If the user has ADMIN priviledges, it will auto-escalate and remember where to run the tools. 
- PreFetch forensics is a neat thing, and Eric Zimmerman has some neat tools for looking at them. However running all the tools in TEA generates over 30 Prefetch Entries, and as Prefetch only can store a finite number of tools you are destroying forensic evidence.
This I move the files from Prefetch to Prefetch.orig before running any tools / memory capture, then back up the files in prefetch.orig to the Tea evidence folder for analysis. Finally I delete the current contents of prefetch (should only have the forensic tools run) and restore the files from prefetch.orig back into prefetch.
I use robocopy to preserve the created and modified dates. While prefetch will continue to add content and delete content, at least the initial files from when you started the process are preserved (this will also allow you to determine which ones were automatically created after)
- There are many memory dump tools out there. In my testing as of 2021-12-14 WinPmem is the fastest and most stable. As of this point I have not tested the captures to see if any particulat tool captures a better dump than the others.
- Encryption should be auto-detected, and if bitlock is mounted the decryption key should be automatically saved out.
- All (almost all? may have missed a few...) outputs from tools as individual files. These are hashed.
- All outputs put in an evidence.txt report, which is also hashed
- verify_report.cmd allows you to drag the folder made from you acquisition, if any files were edited it should provide an error and explain where that error is. This is why it is important not to edit the files.
-  

How can the code be improved?
- All executables moved to a BIN folder, then tested to see if they work. I did this in a former version but some tools had difficulty with relative paths.
- Each piece could be a seperate modular batch file. The config would not just tell that batch to run, but WHEN to run, ie: you could change the order. It would also make editing a script for a specific tool much easier.


Order of volitilaty for TEA
- Backup Prefetch folder
- Capture Memory
- Back up system registry
- Back up user specific registry
- copy backed up prefetch to USB
- Run tools from config
- Detect Encryption (If mounted bitlocker, display recover key to screen AND back up to USB)


Things to consider
- if covert, restore prefetch folder.
- remove USB attachment from registry



Created by:
mark@southby.ca


