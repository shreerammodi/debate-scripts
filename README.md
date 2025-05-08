# Debate Scripts

This is a collection of scripts to extend Verbatim's functionality. The
motivation for creating this repo was to make available the best versions of
scripts I know of in a repo accessible to the community. Additionally, this
enables installation of scripts in a way that doesn't require installation of a
separate version of Verbatim, which can often cause issues with permissions and
antivirus.

Note that the work is not entirely my own and large amounts of inspiration and
certain functions have been taken from other people's implementations of these
macros.

> [!WARNING]
>
> You should generally have an idea of how your computer works before installing
> these scripts. I've comprehensively tested the scripts on my computer, and
> they have been successfully installed on other peoples' computers as well.
> While my instructions are thorough, as a rule of thumb you should understand
> what you're doing to your machine and not blindly follow instructions.

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
left sidebar, expanding Verbatim, and right clicking
`Modules > Insert > Module`. Then, you should copy and paste each script from
the files in this repo into that new module.

You can generally pick and choose which scripts to add, with the exception that
the `helper.vb` script is necessary for the `create-send-doc.vb` and
`create-zap-doc.vb` scripts to work.

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

You should also add the following two styles to your Verbatim template.

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
   - Name: either `Analytic` or `Undertag`
   - Style type: `Linked (paragraph and character)`
   - Style based on: `Tag`, or if you wish, `Normal Text` for `Undertag`
   - Style for following paragraph: `Normal`
4. Aestheticize the style as you wish using the following options:
   - Color, size, italics
5. Check `Add to template`

# Scripts

## `Create Send Doc`

Provides the `CreateSendDoc` macro, which creates a document in your `Downloads`
directory having the same title as the current document with `[S]` prepended.
The send document has all Analytics and Undertags omitted, leaving just
cards/tags/headers.

## `Create Zap Doc`

Provides `Zap`, `CondenseZap`, and `CreateZappedDoc`.

`Zap` deletes all text from card bodies in the current document that is not
highlighted, i.e. any text that is not to be read in a speech.

`CondenseZap` formats the Zapped document properly, removing unnecessary line
breaks in the card bodies due to the way `Zap` works.

`CreateZappedDoc` creates a document in your `Downloads` directory having the
same title as the current document with `[R]` prepended. It then runs `Zap` and
`CondenseZap` on that document.

## `For Reference`

Provides `ForReference`, which operates on a selection of text. It takes all the
highlights in the selected text and turns them Gray. This is useful for
referencing previously read cards in blocks, and for recutting your opponent's
evidence.

## `Highlight to Fill`

Provides `ConvertHighlightsToFills`, which takes all the highlights in a
selection of text and converts them to background fills. This is mainly useful
for recuts of opponents' evidence, to prevent `unihighlight` from standardizing
both your recut and their original highlight. You would first use `ForReference`
on their evidence, then convert it to a fill (to preserve the gray color), and
then rehighlight it.
