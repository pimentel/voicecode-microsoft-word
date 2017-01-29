# Microsoft Word package for VoiceCode

This package provides basic functionality for manipulating text in Microsoft Word with VoiceCode.
Currently, only Microsoft Word Mac is supported.
It has been tested quite a bit with Microsoft Word Mac 2016.

In addition to implementing most of the basic selection commands, `editor:move-to-line-number` (🔉spring🔉) can both jump between lines or pages.
The default is to jump between lines.
This can be changed by the command 🔉page mode🔉.
You can also switch back to line mode with the command 🔉line mode🔉.

## Behavior

Most behavior is as expected, except for the following:

- `editor:expand-selection-to-scope` (🔉bracken🔉): will expand selection to word, sentence, paragraph, section, story

### Issues

Please report any issues to [GitHub issues](https://github.com/pimentel/voicecode-microsoft-word/issues).

### License

MIT License. See `LICENSE` for more info.

### Author

[Harold Pimentel](https://pimentel.github.io)
