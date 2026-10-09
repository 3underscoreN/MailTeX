# MailTeX

MailTeX is an Outlook add-in that allows you to write TeX code and insert the rendered equations directly into your email messages.

> [!NOTE] 
> MailTeX, as well as this README, is currently in early development. Limitations exist, especially for Gmail users. Please read below for more details.

## Features

- Write TeX code right next to your email compose window
- Live preview of the rendered equations
- Insert the equations into your email as PNG images

## Known Issues

- Inserted equations may not display correctly in Gmail on the web or in the Gmail mobile apps.
> [!IMPORTANT]
> Currently, MailTeX inserts the rendered equations as PNG images in the email body. While most email clients support this, **Gmail web mail does not seem to be happy about it**. Therefore, if you are sending emails to Gmail users, please think twice before using this add-in.

## Installation

### Sideloading into Outlook
1. Download [the MailTeX add-in manifest file](https://github.com/3underscoreN/MailTex/blob/main/manifest.xml).
2. Visit the official sideloading website at [https://aka.ms/olksideload](https://aka.ms/olksideload). Click `My add-ins`, then select `Add a custom add-in` from the Custom Addins menu. Choose `Add from File...` and upload the downloaded manifest file.
3. Restart Outlook to activate the add-in.

## Usage

### Outlook Web / New Outlook (Windows)

1. In the Compose menu, click "Apps" on the ribbon, then select "MailTeX" from the list of available add-ins.
2. Type your TeX code in the MailTeX pane.
3. Click the "Insert" button to add the equations to your email as PNG images. Both inline and block equations are supported (we highly suggest using block equations, though).

<details>
<summary>Animated instructions</summary>

![How to use, Outlook Web & New Outlook (Windows)](docs/how-to-use-web.gif)

</details>

### New Outlook (macOS)

1. In the Compose menu. click "..." on the ribbon, then select "MailTeX" from the list.
2. Type your TeX code in the MailTeX pane.
3. Click the "Insert" button to add the equations to your email as PNG images.

<details>
<summary>Animated instructions</summary>

![How to use, New Outlook (macOS)](docs/how-to-use-new.gif)

</details>

## Contributing

Contributions to MailTeX are welcome and greatly appreciated.

Thank you for using MailTeX!

## License

MailTeX is licensed under the MIT License. See [LICENSE](LICENSE) for more details.