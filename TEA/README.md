# TEA
Technological Evidence Acquisition (Proof of Concept Only)

TEA is a Batch file which aims to gather as much information from a Live System in an automated fashion.
TEA was created as a reponse to early versions Microsoft COFEE not working as advertised.
COFEE Issues: 
- It was not open source, could not tell which app was running at which point.
- Early version could not tell the different from Windows XP vs Vista, or 32 bit and 64 bit. TEA worked around these issues in novel ways (would hash and run local tools for vista vs all portable apps for XP)

Why a CMD not Powershell?
- Powershell not always enabled. You want to do minimal changes to a system (Forensic)
- Every Windows machine can run a CMD (BAT is being depreciated in future versions of Windows)

As I can not maintain this tool and batch, it has been modified significantly over the years, and some old code and clean up has not been done, I can not recommend this is used on important cases - but rather used as a reference so you can make a better purpose built tool.
I have shared versions of this tools with other tool developers over the years, however I still find some of their tools to lack the flexibility of a batch file.

Notable features:
- Tamper resistant report generated which can be verified using verify_report.cmd
- Dumps Bitlocker recovery key
- Extracts better registry artifacts thank FTK Imager
- Shows the command line which is run, thus educating anyone using the tool

How to use:
- Configure on your forensic machine using brew.bat
- Copy folder to your Live Search Scene USB
- Run TEA.bat on the target machine
- Once target machine powered off, remove USB, plug back into forensic machine and copy off folder

If you modify this and use it on real files, or it inspired you or a company to make a better first responder tool I would greatly appreciate you letting me know (may be a challenge coin in the deal for you!)

Other things worth mentioning:
- I had started (and abandoned due to shifting priorities) a project which was a hardware device which automated TEA. Think of what a Flipper Zero is today, it was a Teensy Arduino which you could select the target OS, plug it in and the light would blink RED and then turn GREEN when it was ready to be removed.
- I did a Proof of Concept using USB Rubby duckies and it worked as designed. 
- This was to allow the forensic expert to set up some devices and ship them out to front line folk who would just plug it in and gather the info then send it out.
- I wanted to tie it into a tool which would bypass the windows lockscreen, at the time there were very few tools which could do this via DMA. Today I suspect there are many options. 
