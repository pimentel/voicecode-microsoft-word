pack = Packages.register
  name: 'microsoft-word'
  description: 'Functionality for manipulating text in Microsoft Word'
  platforms: [
    'darwin'
  ]
  applications: ['com.microsoft.Word']

runVisualBasic = (whichScript, application = "Microsoft Word") ->
  text = "tell application \"#{application}\" to run VB macro macro name \"#{whichScript}\""
  Actions.applescript text

growlNotification = """
on growlNotification(theMessage)
	tell application "Growl"
		set the allNotificationsList to {"Microsoft Word"}
		set the enabledNotificationsList to {"Microsoft Word"}
		register as application "Microsoft Word" all notifications allNotificationsList default notifications enabledNotificationsList icon of application "Microsoft Word"
		notify with name "Microsoft Word" title "Microsoft Word" description theMessage application name "Microsoft Word"
	end tell
end growlNotification
"""

pack.settings =
  'moveMode': 'line'

pack.command 'add-citation',
  spoken: 'add citation'
  description: "add citation using Zotero. Requires Zotero Microsoft Word plug-in."
  action: ->
    runVisualBasic 'ZoteroInsertCitation'
pack.command 'change-move-mode-page',
  spoken: 'page mode'
  description: '"jump" by page in Microsoft Word. Used with editor:move-to-line-number.'
  action: ->
    pack.settings.moveMode = 'page'
pack.command 'change-move-mode-line',
  spoken: 'line mode'
  description: '"jump" by page in Microsoft Word. Used with editor:move-to-line-number.'
  action: ->
    pack.settings.moveMode = 'line'
pack.command 'toggle-line-numbering',
  spoken: 'toggle numbers'
  description: 'Toggle line numbering. NOTE: makes line numbering continuous which may change your page formatting.'
  action: ->
    cmd = """
    tell application "Microsoft Word"
    	set pageSetup to page setup of active document
    	set lineNumbering to line numbering of pageSetup
    	set isActive to active line of lineNumbering
    	set active line of lineNumbering to not isActive
    	set restart mode of lineNumbering to restart continuous
    end tell
    """
    @applescript cmd
pack.command 'word-count',
  spoken: 'word count'
  description: "Display a word and character count using Growl."
  action: ->
    cmd = """
    #{growlNotification}

    tell application "Microsoft Word"
    	set numberOfCharacters to compute statistics of active document statistic statistic characters with spaces
    	set numberOfWords to compute statistics of active document statistic statistic words
      set theNotification to "chars: " & numberOfCharacters & "\nwords: " & numberOfWords
      -- can also use the below if you don't want to use Growl
    	-- display alert theNotification
      my growlNotification(theNotification)
    end tell
    """
    @applescript cmd

pack.implement
  'editor:move-to-line-number': (input) ->
    cmd = """
    tell application "Microsoft Word"
    	set textObject to text object of selection
    	navigate textObject position absolute count #{input} to goto a #{pack.settings.moveMode} item
    end tell
    """
    @applescript cmd
  'editor:move-to-line-number-and-way-right': (input) ->
    cmd = """
    tell application "Microsoft Word"
    	set textObject to text object of selection
    	navigate textObject position absolute count #{input} to goto a line item
    end tell
    """
    @applescript cmd
    @do 'cursor:way-right'
  'cursor:way-down': ->
    runVisualBasic 'EndOfDocument'
  'cursor:way-up': ->
    runVisualBasic 'StartOfDocument'
  'common:zoom-in': ->
    runVisualBasic 'ViewZoomIn'
  'common:zoom-out': ->
    runVisualBasic 'ViewZoomOut'
  'common:indent-right': ->
    runVisualBasic 'IncreaseIndent'
  'common:indent-left': ->
    runVisualBasic 'DecreaseIndent'
  'delete:lines': ({first, last} = {}) ->
    @do 'cursor:way-right'
    @do 'selection:way-left'
    @do 'common:delete'
  'selection:next-occurrence': (input) ->
    cmd = """
    tell application "Microsoft Word"
    	tell find object of selection
    		execute find find text "#{input.value}" with match forward
    	end tell
    end tell
    """
    @applescript cmd
  'selection:extend-to-next-occurrence': (input) ->
    cmd = """
    tell application "Microsoft Word"
      set extend mode of selection to true
    	tell find object of selection
    		execute find find text "#{input.value}" with match forward
    	end tell
      set extend mode of selection to false
    end tell
    """
    @applescript cmd

  'selection:previous-occurrence': (input) ->
    cmd = """
    tell application "Microsoft Word"
    	tell find object of selection
    		execute find find text "#{input.value}" without match forward
    	end tell
    end tell
    """
    @applescript cmd
  'selection:extend-to-previous-occurrence': (input) ->
    cmd = """
    tell application "Microsoft Word"
      set extend mode of selection to true
    	tell find object of selection
    		execute find find text "#{input.value}" without match forward
    	end tell
      set extend mode of selection to false
    end tell
    """
    @applescript cmd

  'clipboard:paste-from-history': ->
    cmd = """
    tell application "Microsoft Word"
    	set textObject to text object of selection
    	paste and format textObject type format plain text
    end tell
    """
    @applescript cmd

  'selection:next-word': (input) ->
    if not input?
      cmd = """
      tell application "Microsoft Word"
        -- clear the selection
      	set selection start of selection to selection end of selection
      	set startOfSelection to selection start of selection
      	set textObject to text object of selection
      	set newPosition to move end of range textObject by a word item
      	set selection start of selection to startOfSelection
      end tell
      """
      @applescript cmd
    else
      input = input - 1
      cmd = """
      tell application "Microsoft Word"
        -- clear the selection
      	set selection start of selection to selection end of selection
        set textObject to text object of selection
      	set newPosition to move start of range textObject by a word item count #{input}
        set theResult to expand selection by a word item
      end tell
      """
      @applescript cmd
  'selection:previous-word': (input) ->
    if not input?
      cmd = """
      tell application "Microsoft Word"
      	-- clear the selection
      	set selection end of selection to selection start of selection
      	set startOfSelection to selection start of selection
      	set textObject to text object of selection
      	set newPosition to move end of range textObject by a word item count -1
      	set selection end of selection to startOfSelection
      end tell
      """
      @applescript cmd
    else
      input = -input
      cmd = """
      tell application "Microsoft Word"
        -- clear the selection
        set selection end of selection to selection start of selection
        set textObject to text object of selection
      	set newPosition to move end of range textObject by a word item count #{input}
        set theResult to expand selection by a word item
      end tell
      """
      @applescript cmd

  'selection:way-up': ->
    cmd = """
    tell application "Microsoft Word"
    	set selection start of selection to 0
    end tell
    """
    @applescript cmd
  'selection:way-down': ->
    cmd = """
    tell application "Microsoft Word"
    	set storyRange to get story range of (document of active window) story type main text story
    	set selection end of selection to end of content of storyRange
    end tell
    """
    @applescript cmd

  'selection:next-selection-occurrence': ->
    cmd = """
    tell application "Microsoft Word"
    	set selectionType to selection type of selection
    	if selectionType is selection normal then
    		tell find object of selection
    			execute find find text content of selection with match forward
    		end tell
    	end if
    end tell
    """
    @applescript cmd
  'selection:previous-selection-occurrence': ->
    cmd = """
    tell application "Microsoft Word"
    	set selectionType to selection type of selection
    	if selectionType is selection normal then
    		tell find object of selection
    			execute find find text content of selection without match forward
    		end tell
    	end if
    end tell
    """
    @applescript cmd
