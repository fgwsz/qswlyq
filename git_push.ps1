git add ./*.py
git add ./file_template.docx
git add ./*.ps1
$commit_info=Read-Host -Prompt "Please input commit info"
git commit -m $commit_info
git push
