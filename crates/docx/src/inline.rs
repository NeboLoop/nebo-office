use nebo_spec::{Run, TextRun};

/// Parse markdown-like inline text into a Vec of Runs.
/// Supported: **bold**, *italic*, __underline__, ~~strike~~, `code`, [text](url), [text](#bookmark), [^1]
pub fn parse_inline_text(text: &str) -> Vec<Run> {
    let mut runs = Vec::new();
    let chars: Vec<char> = text.chars().collect();
    let mut pos = 0;
    let mut current_text = String::new();

    while pos < chars.len() {
        let ch = chars[pos];

        // Footnote reference: [^1]
        if ch == '[' && pos + 1 < chars.len() && chars[pos + 1] == '^' {
            if let Some(end) = find_char(&chars, pos + 2, ']') {
                flush_text(&mut current_text, &mut runs);
                let id: String = chars[pos + 2..end].iter().collect();
                runs.push(Run::FootnoteRef(nebo_spec::FootnoteRun { footnote: id }));
                pos = end + 1;
                continue;
            }
        }

        // Link: [text](url) or [text](#bookmark)
        if ch == '[' {
            if let Some(text_end) = find_char(&chars, pos + 1, ']') {
                if text_end + 1 < chars.len() && chars[text_end + 1] == '(' {
                    if let Some(url_end) = find_char(&chars, text_end + 2, ')') {
                        flush_text(&mut current_text, &mut runs);
                        let link_text: String = chars[pos + 1..text_end].iter().collect();
                        let url: String = chars[text_end + 2..url_end].iter().collect();
                        runs.push(Run::Text(TextRun {
                            text: link_text,
                            link: Some(url),
                            bold: None,
                            italic: None,
                            underline: None,
                            strike: None,
                            superscript: None,
                            subscript: None,
                            font: None,
                            size: None,
                            color: None,
                            highlight: None,
                            all_caps: None,
                            small_caps: None,
                        }));
                        pos = url_end + 1;
                        continue;
                    }
                }
            }
        }

        // Bold: **text**
        if ch == '*' && pos + 1 < chars.len() && chars[pos + 1] == '*' {
            if let Some(end) = find_double(&chars, pos + 2, '*') {
                flush_text(&mut current_text, &mut runs);
                let inner: String = chars[pos + 2..end].iter().collect();
                runs.push(Run::Text(TextRun {
                    text: inner,
                    bold: Some(true),
                    italic: None,
                    underline: None,
                    strike: None,
                    superscript: None,
                    subscript: None,
                    font: None,
                    size: None,
                    color: None,
                    highlight: None,
                    link: None,
                    all_caps: None,
                    small_caps: None,
                }));
                pos = end + 2;
                continue;
            }
        }

        // Italic: *text*
        if ch == '*' && (pos + 1 >= chars.len() || chars[pos + 1] != '*') {
            if let Some(end) = find_single_not_double(&chars, pos + 1, '*') {
                flush_text(&mut current_text, &mut runs);
                let inner: String = chars[pos + 1..end].iter().collect();
                runs.push(Run::Text(TextRun {
                    text: inner,
                    italic: Some(true),
                    bold: None,
                    underline: None,
                    strike: None,
                    superscript: None,
                    subscript: None,
                    font: None,
                    size: None,
                    color: None,
                    highlight: None,
                    link: None,
                    all_caps: None,
                    small_caps: None,
                }));
                pos = end + 1;
                continue;
            }
        }

        // Underline: __text__
        if ch == '_' && pos + 1 < chars.len() && chars[pos + 1] == '_' {
            if let Some(end) = find_double(&chars, pos + 2, '_') {
                flush_text(&mut current_text, &mut runs);
                let inner: String = chars[pos + 2..end].iter().collect();
                runs.push(Run::Text(TextRun {
                    text: inner,
                    underline: Some(true),
                    bold: None,
                    italic: None,
                    strike: None,
                    superscript: None,
                    subscript: None,
                    font: None,
                    size: None,
                    color: None,
                    highlight: None,
                    link: None,
                    all_caps: None,
                    small_caps: None,
                }));
                pos = end + 2;
                continue;
            }
        }

        // Strikethrough: ~~text~~
        if ch == '~' && pos + 1 < chars.len() && chars[pos + 1] == '~' {
            if let Some(end) = find_double(&chars, pos + 2, '~') {
                flush_text(&mut current_text, &mut runs);
                let inner: String = chars[pos + 2..end].iter().collect();
                runs.push(Run::Text(TextRun {
                    text: inner,
                    strike: Some(true),
                    bold: None,
                    italic: None,
                    underline: None,
                    superscript: None,
                    subscript: None,
                    font: None,
                    size: None,
                    color: None,
                    highlight: None,
                    link: None,
                    all_caps: None,
                    small_caps: None,
                }));
                pos = end + 2;
                continue;
            }
        }

        // Code: `text`
        if ch == '`' {
            if let Some(end) = find_char(&chars, pos + 1, '`') {
                flush_text(&mut current_text, &mut runs);
                let inner: String = chars[pos + 1..end].iter().collect();
                runs.push(Run::Text(TextRun {
                    text: inner,
                    font: Some("Courier New".into()),
                    bold: None,
                    italic: None,
                    underline: None,
                    strike: None,
                    superscript: None,
                    subscript: None,
                    size: None,
                    color: None,
                    highlight: None,
                    link: None,
                    all_caps: None,
                    small_caps: None,
                }));
                pos = end + 1;
                continue;
            }
        }

        current_text.push(ch);
        pos += 1;
    }

    flush_text(&mut current_text, &mut runs);
    runs
}

fn flush_text(current: &mut String, runs: &mut Vec<Run>) {
    if !current.is_empty() {
        runs.push(Run::Text(TextRun {
            text: std::mem::take(current),
            bold: None,
            italic: None,
            underline: None,
            strike: None,
            superscript: None,
            subscript: None,
            font: None,
            size: None,
            color: None,
            highlight: None,
            link: None,
            all_caps: None,
            small_caps: None,
        }));
    }
}

fn find_char(chars: &[char], start: usize, target: char) -> Option<usize> {
    for i in start..chars.len() {
        if chars[i] == target {
            return Some(i);
        }
    }
    None
}

fn find_double(chars: &[char], start: usize, target: char) -> Option<usize> {
    for i in start..chars.len() - 1 {
        if chars[i] == target && chars[i + 1] == target {
            return Some(i);
        }
    }
    None
}

fn find_single_not_double(chars: &[char], start: usize, target: char) -> Option<usize> {
    for i in start..chars.len() {
        if chars[i] == target {
            if i + 1 < chars.len() && chars[i + 1] == target {
                continue;
            }
            return Some(i);
        }
    }
    None
}

/// Convert runs back to markdown-like string if possible.
/// Returns None if the runs contain formatting that can't be represented in markdown.
pub fn runs_to_markdown(runs: &[Run]) -> Option<String> {
    let mut result = String::new();

    for run in runs {
        match run {
            Run::Text(tr) => {
                let has_complex = tr.font.is_some()
                    || tr.size.is_some()
                    || tr.color.is_some()
                    || tr.highlight.is_some()
                    || tr.superscript == Some(true)
                    || tr.subscript == Some(true)
                    || tr.all_caps == Some(true)
                    || tr.small_caps == Some(true);

                if has_complex && !is_only_link(tr) {
                    return None;
                }

                if let Some(url) = &tr.link {
                    result.push_str(&format!("[{}]({})", tr.text, url));
                } else if tr.bold == Some(true) {
                    result.push_str(&format!("**{}**", tr.text));
                } else if tr.italic == Some(true) {
                    result.push_str(&format!("*{}*", tr.text));
                } else if tr.underline == Some(true) {
                    result.push_str(&format!("__{}__", tr.text));
                } else if tr.strike == Some(true) {
                    result.push_str(&format!("~~{}~~", tr.text));
                } else if tr.font.as_deref() == Some("Courier New") {
                    result.push_str(&format!("`{}`", tr.text));
                } else {
                    result.push_str(&tr.text);
                }
            }
            Run::FootnoteRef(fr) => {
                result.push_str(&format!("[^{}]", fr.footnote));
            }
            Run::Tab(_) | Run::Field(_) | Run::Break(_) => return None,
            Run::Delete(_) | Run::Insert(_) => return None,
            Run::CommentStart(_) | Run::CommentEnd(_) => return None,
        }
    }

    Some(result)
}

fn is_only_link(tr: &TextRun) -> bool {
    tr.link.is_some()
        && tr.bold.is_none()
        && tr.italic.is_none()
        && tr.underline.is_none()
        && tr.strike.is_none()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_bold() {
        let runs = parse_inline_text("Hello **world**!");
        assert_eq!(runs.len(), 3);
    }

    #[test]
    fn test_italic() {
        let runs = parse_inline_text("Hello *world*!");
        assert_eq!(runs.len(), 3);
    }

    #[test]
    fn test_link() {
        let runs = parse_inline_text("See [details](https://example.com).");
        assert_eq!(runs.len(), 3);
    }

    #[test]
    fn test_footnote() {
        let runs = parse_inline_text("Note[^1].");
        assert_eq!(runs.len(), 3);
    }

    #[test]
    fn test_round_trip() {
        let text = "Hello **bold** and *italic* and __underline__";
        let runs = parse_inline_text(text);
        let md = runs_to_markdown(&runs).unwrap();
        assert_eq!(md, text);
    }
}
