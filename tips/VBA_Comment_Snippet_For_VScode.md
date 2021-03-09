# How to add comment snippet to vs Code
1. Add snippet to vscode
  - Open vs cdoe
  - press ctrl + shift + p 
  - select "preference: Configure user preference"
  - Select "New Global Snippets file"
3. Copy paste the following code to that json file

```json
{
	// Place your snippets for vba here. Each snippet is defined under a snippet name and has a prefix, body and 
	// description. The prefix is what is used to trigger the snippet and the body will be expanded and inserted. Possible variables are:
	// $1, $2 for tab stops, $0 for the final cursor position, and ${1:label}, ${2:another} for placeholders. Placeholders with the 
	// same ids are connected.
	// Example:
	"Print to console": {
		"prefix": "comment",
		"body": [
			"'/**",
			"' * @Purpose: ",
			"' * @Param  : {}",
			"' * @Return : {}",
			"' */"
		],
		"description": "Log output to console"
	}
}

````

3. Type "comment" to .bas file (you will see snippet)
