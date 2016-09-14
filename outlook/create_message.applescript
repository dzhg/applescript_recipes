-- This script is tested with Microsoft Output for Mac Version 15.25 on Mac OS El Capitan
tell application "Microsoft Outlook"
	set theContent to "theContent"
	set theSubject to "theSubject"
	set emails to {{address:"name1@test.com", name:"name1"}, {address:"name2@test.com", name:"name2"}}
	set theMessage to make new outgoing message with properties {subject:theSubject, content:theContent}
	repeat with email in emails
		set addr to address of email
		set n to name of email
		make new recipient with properties {email address:{address:addr, name:n}} at end of to recipients of theMessage
	end repeat
end tell
