#!/bin/bash
# set -x
FILE="$1/Office2019.pkg"
SERIAL="$1/Microsoft_Office_2019_VL_Serializer.pkg"
ACTION=$2
# Define first run configure function
ConfigureOfficeFirstRun() {
    # This function will configure the first run dialog windows for all Office 2016 apps.
    # It will also set the desired diagnostic info settings for Office application.
    
    # Special check for OneNote as the application name and PLIST name are not the same.
    if [[ $app == OneNote ]]
    then
        app="onenote.mac";
    fi
   defaults write /Library/Preferences/com.microsoft."$app" kSubUIAppCompletedFirstRunSetup1507 -bool true
   defaults write /Library/Preferences/com.microsoft."$app" SendAllTelemetryEnabled -bool false
   defaults write Library/Preferences/com.microsoft."$app"  kFREEnterpriseTelemetryInfoKey -bool TRUE
   defaults write Library/Preferences/com.microsoft."$app"  kFRETelemetryConsentKey -bool TRUE
   defaults write Library/Preferences/com.microsoft."$app"  SendCrashReportsEvenWithTelemetryDisabled -bool FALSE
   defaults write Library/Preferences/com.microsoft."$app"  SendASmileEnabled -bool false
    # Outlook and OneNote require one additional first run setting to be disabled
    if [[ $app == "Outlook" ]] || [[ $app == "onenote.mac" ]]; then
        defaults write /Library/Preferences/com.microsoft."$app" FirstRunExperienceCompletedO15 -bool true
    fi
}

# should download the latest version to current working dir
download_office() {
	curl https://go.microsoft.com/fwlink/?linkid=525133 -L -o "$FILE"
}

# just use a dummy, test file, for script debugging
download_test() {
	curl http://ipv4.download.thinkbroadband.com/10MB.zip -L -o "$FILE"
	sleep 5
}

# install all options for package
install_office() {
	installer -pkg "$FILE" -target /
}

# don't actually do anything, except pause to display any relevant messages
install_test() {
	sleep 8
}

# Activate
serialiser() {
	installer -pkg "$SERIAL" -target /
}

# Set some defaults
config_options() {
	# Configure Office First Run behaviour now
	OfficeApps=(Excel OneNote Outlook PowerPoint Word)
	for APPNAME in ${OfficeApps[*]}
	do
	    if [[ -e "/Applications/Microsoft $APPNAME.app" ]]; then
	    app=$APPNAME
	    ConfigureOfficeFirstRun
	    fi
	done
	# Configure AutoUpdate behaviour (set to manual check and hide insider program)
	defaults write /Library/Preferences/com.microsoft.autoupdate2 HowToCheck -string 'AutomaticDownload'
	defaults write /Library/Preferences/com.microsoft.autoupdate2 LastUpdate -date '2021-02-11T15:00:00Z'
	defaults write /Library/Preferences/com.microsoft.autoupdate2 DisableInsiderCheckbox -bool true
	defaults write /Library/Preferences/com.microsoft.autoupdate2 SendAllTelemetryEnabled -bool FALSE
	defaults write /Library/Preferences/com.microsoft.autoupdate.fba SendAllTelemetryEnabled -bool FALSE
	# Configure the default save location for Office for all existing users
	defaults write /Library/Preferences/com.microsoft.office DefaultsToLocalOpenSave -bool true
	defaults write /Library/Preferences/com.microsoft.office TermsAccepted1809 -bool TRUE
	defaults write /Library/Preferences/com.microsoft.office ConnectedOfficeExperiencesPreference
	# this one probably requires a configuration profile, see macadmins/preferences
	defaults write /Library/Preferences/com.microsoft.office OptionalConnectedExperiencesPreference -bool false
	# Error Reporting
	defaults write Library/Preferences/com.microsoft.errorreporting SendCrashReportsEvenWithTelemetryDisabled -bool TRUE
	# OFFICE365SERCICEV2
	defaults write Library/Preferences/com.microsoft.Office365ServiceV2 SendAllTelemetryEnabled -bool FALSE
	# disable RequiredData Notice alert
	# https://github.com/rtrouton/rtrouton_scripts/tree/master/rtrouton_scripts/disable_mau_required_data_notice_screen
	# This script is designed to suppress the Microsoft AutoUpdate Required Data Notice screen
	# The script runs the following actions:
	# 
	# 1. Identifies all users on the Mac with a UID greater than 500
	# 2. Identifies the home folder location of all users identified
	#    in the previous step.
	# 3. Sets the com.microsoft.autoupdate2.plist file with the following
	#    key and value. This will suppress Microsoft AutoUpdate\'s 
	#    Required Data Notice screen and stop it from appearing.
	#
	#    Key: AcknowledgedDataCollectionPolicy
	#    Value: RequiredDataOnly
	# Identify all users on the Mac with a UID greater than 500
	allLocalUsers=$(dscl . -list /Users UniqueID | awk '$2>500 {print $1}')

	for userName in ${allLocalUsers}; do

		  # Identify the home folder location of all users with a UID greater than 500.
		  # this only works for local users, not domain users

		  userHome=$(dscl . -read "/Users/$userName" NFSHomeDirectory 2>/dev/null | sed 's/^[^\/]*//g')
	  
		  # Verify that home folder actually exists.
	  
		  if [[ -d  "$userHome" ]]; then

			# If the home folder exists, sets the com.microsoft.autoupdate2.plist file with the needed key and value.

			defaults write "${userHome}/Library/Preferences/com.microsoft.autoupdate2.plist" AcknowledgedDataCollectionPolicy -string 'RequiredDataOnly'

			# This script is designed to be run with root privileges, so the ownership of the com.microsoft.autoupdate2.plist file
			# and the enclosing directories are re-set to that of the account which owns the home folder. 

			chown "$userName" "${userHome}/Library/"
			chown "$userName" "${userHome}/Library/Preferences"
			chown "$userName" "${userHome}/Library/Preferences/com.microsoft.autoupdate2.plist"
	  
		  fi

	done
	# exit 0
}

case "$2" in
	download)
		download_office
		;;
	setup)
		install_office
		;;
	activate)
		serialiser
		;;
	config)
		config_options
		;;
	download_test)
		download_test
		;;
	setup_test)
		install_test
		;;
	*)
		download_office
		install_office
		serialiser
		config_options
		;;
esac
