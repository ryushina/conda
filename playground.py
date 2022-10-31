import re

is_with_parentheses = False
is_to_be_trimmed = False
txt = "(M. MAURICIO)"
fixedMiddleName = ''
withParenthesesPattern = "^\(.*\)$"
withMItoTrim = "^\(.\."
withParentheses = re.search(withParenthesesPattern, txt).string
if withParentheses is not None:
    is_with_parentheses = True

if is_with_parentheses == True:
    MIToTrim = re.findall(withMItoTrim, txt)
    if len(MIToTrim) != 0:
        fixedMiddleName = txt[1:3]
    else:
        fixedMiddleName = txt[1:-1]
print(fixedMiddleName)
