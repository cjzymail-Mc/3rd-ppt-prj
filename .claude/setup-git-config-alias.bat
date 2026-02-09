@echo off
echo Setting up Git global configurations...

git config --global user.name "Mc"
git config --global user.email "cjzymail@gmail.com"
git config --global alias.st status
git config --global alias.chk checkout
git config --global alias.sw switch
git config --global alias.br branch
git config --global alias.ls "ls-tree -r --name-only"
git config --global core.quotePath false

echo Configuration complete! Run 'git config --global --list' to verify.
pause