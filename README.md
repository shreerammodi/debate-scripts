# Debate Scripts

This is a collection of scripts to extend Verbatim's functionality and provide
macros that are useful in competitive debate.

Note that the work is not entirely my own and large amounts of inspiration and
certain functions have been taken from other people's implementations of these
macros.

# Installation

## Module Setup

To install scripts in Word you need to add them as a Module in the Visual Basic
section of the existing MS Word template that you want them to be available in.
For Verbatim scripts, this is the `Debate.dotm` template.

### Mac

> Go to the menu bar > Tools > Macro > Visual Basic Editor

### Windows

> Go to the Word home menu > options (bottom) > customize ribbon > in
> `Main Tabs` check `developer`

You'll now have a developer tab in the Word ribbon. Then,

> Go to Developer > Visual Basic

## Adding Macros

I would recommend adding macros to a new module. You can do this by going to the
sidebar on the left, expanding Verbatim, right clicking, then going into
`Modules > Insert > Module`. Then, you should copy and paste each script from
the files in this repo into that new module.

You can generally pick and choose which scripts to add, with the exception that
the `helper.vb` script is necessary for the `send-doc` and `zapper` scripts to
work.

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

You _must_ add the following two styles to your Verbatim template.

- Analytic: for analytic arguments that you want omitted from the doc you send
  to your opponent.
- Undertag: for brief notes that go below tags which you want omitted from the
  doc you send to your opponent.

> [!NOTE]
>
> Some teams have their style for Analytic arguments named as `Analytics`. In my
> opinion, this is incorrect. All of the other styles that Verbatim provides are
> in the singular tense, e.g. `Tag`, `Cite`, `Hat`, etc. and not in the plural
> tense, e.g. `Tags`, `Cites`, etc. Therefore, to maintain consistency with the
> default Verbatim styles, any added styles should also be in the singular
> tense.

To add styles to your template:

1. Open the styles pane
2. Click `New Style`
3. Use the following settings:
   - Name: `Analytic` / `Undertag`
   - Style type: `Linked (paragraph and character)`
   - Style based on: `Tag` for `Analytic` and `Normal Text` for `Undertag`
   - Style for following paragraph: `Normal`
4. Aestheticize the style as you wish using the following options:
   - Color, size, italics
5. Check `Add to template`

# Scripts

We provide the following scripts:

- **Zapper**: this deletes all text which is not to be read in round, like
  non-highlighted card body. There are two callable functions:
  - `CreateZappedDoc`: this creates a new document which is a zapped version of
    the currently opened doc.
  - `ZapCard`: with your cursor in the tag of a card, calling this zaps only the
    current card in the current document.
- `CreateSendDoc`: Given a speech document, creates a new send doc without
  analytics and undertags.
- `ForReference`: Changes the highlight color of selected text to gray and
  converts it to a background fill, preventing it from being affected by
  standardize highlighting.
- **Card Marker**: Allows you to conveniently mark a card at your cursor and
  then compile all the marked cards in the document at the end for easy
  reference.
  - `MarkCard`: creates a "mark" at the cursor's position in the body of a card,
    shading the rest of the card body red to indicate it wasn't read during the
    speech.
  - `CompileMarkedCards`: creates a section at the bottom of the doc with all
    the cards marked in the current document.
