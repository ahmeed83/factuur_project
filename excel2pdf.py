import subprocess


def convert_excel_to_pdf(name):
    cmd = f"""
        on run
                set thisFile to ((path to desktop as text) & "{name}") as alias
                tell application "Finder"
                    set theItemParentPath to container of (thisFile as alias) as text
                    set theItemName to (name of thisFile) as string
                    set theItemExtension to (name extension of thisFile)
                    set theItemExtensionLength to (count theItemExtension) + 1
                    set theOutputPath to theItemParentPath & (text 1 thru (-1 - theItemExtensionLength) of theItemName)
                    set theOutputPath to (theOutputPath & ".pdf")
                end tell
                tell application "Microsoft Excel"
                    set isRun to running
                    activate
                    open thisFile
                    tell active workbook
                        alias theOutputPath
                        -- set overwrite to true
                        save workbook as filename theOutputPath file format PDF file format with overwrite
                        --save overwrite yes
                        close saving no
                    end tell
                    -- close active workbook saving no
                    if not isRun then quit
                end tell
        end run
    """
    result = subprocess.run(['osascript', '-e', cmd], capture_output=True)
    return result.stdout
