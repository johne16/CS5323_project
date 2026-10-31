# main research question
- when scanning a qr code, the link an app shows you can differ from the site it actually opens

# methods
- they tested 126 crafted qr codes through different apps (browsers, qr scanners, and camera apps on Android and iOS)
- Android testing was semi-automated
- iOS and Google Lens were tested manually

# main findings
see findings section in [5. Findings](./Finding_and_Exploiting_Discrepancies_in_QR_Code_Payload_Processing_Summary.md)

# main contributions
- Categorized the ways apps process and display QR code payloads differently.
- Tested popular mobile browsers and QR scanners for mismatches that enable deception.
- Built a semi-automated Android framework to find exploitable mismatches and narrow down their root causes.
- Analyzed the mismatches, how they can be exploited, and how to mitigate them.

# novelty of the research
- the paper introduces three stages in the consumption of QR codes:
1. decoding
2. optional parsing, validation and rendering
3. execution
-previous work focuses on stages 1 and 3. this paper focuses on stage 2: parsing and display of the QR code's content and how that differs from what it opens across apps and operating systems

# experimental design and quality strengths
- End-to-end testing: they scanned QR codes in the real apps instead of testing URL parsers in isolation.
- Android measurement: they captured all three stages, using Frida for decoded data, screenshots and OCR for the display, and mitmproxy for the domain visited.
- Display checks: OCR mismatches were checked by hand.
- Test cases: they were drawn from earlier attacks and bug reports, which suits visual deception that fuzzing struggles to detect.
- Reproducibility: the code, test cases, and QR images are public.

# weaknesses / limitations
- No user study: they show that displays can mislead but don't measure whether people are actually fooled.
- Uneven testing: iOS apps and Google Lens were tested by hand, and the authors couldn't capture the raw text those apps decoded from each QR code.
  - Google Lens: the Android app's code is heavily obfuscated (deliberately scrambled), so they couldn't find the right method to hook with Frida.
  - iOS: Xcode's iOS simulators don't support virtual cameras, and they couldn't jailbreak their iOS 17.5.1 test device. They also didn't have access to Corellium, a virtual iOS device platform.
- Limited scope: they tested a selection of popular apps, only two secure scanner vendors, and only URL payloads. They excluded retail and e-commerce apps, and the 126 targeted cases can't cover every mismatch.
- Root causes: some are partly speculative because the apps are closed source, such as Trend Micro's scanner.

# Their recommendations
Their recommendation is that apps process each QR code URL in one consistent way:
1. Decode the payload with a set encoding, rejecting invalid bytes, NULL bytes, and ambiguous control characters.
2. Parse it with a WHATWG parser to get one normalized URL.
3. Use that URL for the safety check and confirmation screen, showing the full host.
4. After the user confirms, open that same URL.

They also recommend following Chrome's IDN rules for when to show Punycode, and agreeing on WHATWG as the standard for QR code URLs.

# personal reflections
## do i agree with their conclusions?
- I agree with their conclusions.
- Adopting a published standard is a good way to address vulnerabilities caused by companies processing QR code URLs in their own ways.
- Their mitigation is a good starting point for a general QR code standard.

# connections with other readings/theories


# future research questions
- Other delivery methods: test NFC tags, SMS link previews, and other barcode types such as Data Matrix and Aztec for the same mismatches.
- User study: measure how often people are fooled by each trick, and which display formats or warnings help.
- Other payload types: test Wi-Fi, contact, calendar, and binary payloads for the same mismatches.

# new knowledge/inspiration
- i learned that there is no current QR code handling standard which makes this attack vector sound surprisingingly easy to take advantage of. i'm disappointed that there isn't a required security standard already in place