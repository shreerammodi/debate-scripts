# Debate Scripts

This is a collection of scripts to extend Verbatim's functionality and provide
macros that are useful in competitive debate.

Note that the work is not entirely my own and large amounts of inspiration and
certain functions have been taken from other people's implementations of these
macros.

## Installation

### Module Setup

To install scripts in Word you need to add them as a Module in the Visual Basic
section of the existing MS Word template that you want them to be available in.
For Verbatim scripts, this is the `Debate.dotm` template.

#### Mac

> Go to the menu bar > Tools > Macro > Visual Basic Editor

#### Windows

> Go to the Word home menu > options (bottom) > customize ribbon > in
> `Main Tabs` check `Developer`

You'll now have a developer tab in the Word ribbon. Then,

> Go to Developer > Visual Basic

### Adding Macros

For convenience, you can add all the macros to the `Custom` module that Verbatim
provides. You can access this by going into the sidebar on the left, expanding
`Verbatim`, and clicking on the `Custom` module. Then, you should copy and paste
each script from the files in this repository into the custom module.

> [!IMPORTANT]
>
> You can generally pick and choose which scripts to add, with the exception
> that the `helper.vb` script is always required.

## Adding Keybindings

After adding the macros, you also need to add keybindings to trigger them.

### Mac

1. Go to the menu bar > Tools > Customize Keyboard
2. Make sure you set `Save changes in` to `Debate.dotm`
3. In `Categories` scroll down and click `Macros` then find the macro you want
   to bind
4. Add the keyboard shortcut, then click `Save`

### Windows

To do this in Windows, follow
[this guide](https://support.microsoft.com/en-us/office/customize-keyboard-shortcuts-9a92343e-a781-4d5a-92f1-0f32e3ba5b4d).

## Adding Styles

You _should_ add the following two styles to your Verbatim template.

- Analytic[^1]: for analytic arguments that you want omitted from the doc you
  send to your opponent.
- Undertag: for brief notes that go below tags which you want omitted from the
  doc you send to your opponent and also don't want to show up in the sidebar.

To add styles to your template, do the following in some verbatimized document:

1. Open the styles pane
2. Click `New Style`
3. Name it `Analytic` / `Undertag`
4. Check `Add to template`

For `Analytic`, go back to your document, select a `Tag`, right-click the
`Analytic` style in your sidebar, and click `Update to match selection`.

You can aestheticize the styles as you wish.

## Scripts

We provide the following scripts:

- **Zapper**: this deletes all text which is not to be read in round, like
  non-highlighted card body. There are two callable functions:
  - `CreateZappedDoc`: this creates a new document which is a zapped version of
    the currently opened doc.
  - `ZapCard`: with your cursor in the tag of a card, calling this zaps only the
    current card in the current document.
- `CreateSendDoc`: Given a speech document, creates a new send doc without
  analytics and undertags.
- **ForReference**: changes the highlight color of selected text to gray and
  converts it to a background fill, preventing it from being affected by
  standardize highlighting. Also optionally shrinks text size down 2 font sizes.
  - `ForReferenceFast`: more performant version, which sometimes may remove
    emphasis
  - `ForReferenceSlow`: does not remove any styles, but less performant
- **Card Marker**: Allows you to conveniently mark a card at your cursor and
  then compile all the marked cards in the document at the end for easy
  reference.
  - `MarkCard`: creates a "mark" at the cursor's position in the body of a card,
    shading the rest of the card body red to indicate it wasn't read during the
    speech.
  - `CompileMarkedCards`: creates a section at the bottom of the doc with all
    the cards marked in the current document.

[^1]:
    Some teams have their style for Analytic arguments named as `Analytics`. In
    my opinion, this is incorrect. All of the other styles that Verbatim
    provides are in the singular tense, e.g. `Tag`, `Cite`, `Hat`, etc. and not
    in the plural tense, e.g. `Tags`, `Cites`, etc. Therefore, to maintain
    consistency with the default Verbatim styles, any added styles should also
    be in the singular tense.
