# Office2019macOS
installer for Office 2019

copy this file and Microsoft_Office_2019_VL_Serializer.pkg to the same dir
then call this with e.g. 
Launch $(KACE_DEPENDENCY_DIR)\Office2019macOS.sh 
with params "$(KACE_DEPENDENCY_DIR)" download

parameters download_test & setup_test may save time in debug
also see set -x for more debug output
in debug call from terminal e.g.

sudo ./Office2019macOS.sh /Library/Application\ Support/Quest/KACE/data/kbots_cache/packages/kbots/853 config