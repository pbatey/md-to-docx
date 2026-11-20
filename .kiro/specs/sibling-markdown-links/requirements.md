# Requirements Document

## Introduction

This feature lets a tree of Markdown files be converted to a tree of Word documents whose links between files keep working. Today a Markdown link such as `[Setup](setup.md)` is rendered in the .docx as plain text, not as a hyperlink, so the reader cannot click through to the converted document. With this feature, a link to another local Markdown file is written as a relative hyperlink to the corresponding .docx file. The .docx files can then be opened from one another when they are stored in the same relative layout as the source Markdown files.

By default the hyperlinks are relative, which works for local folders and synced libraries. Readers who open the .docx files in a browser from SharePoint cannot follow relative links, because Word for the web does not resolve relative paths. For that case, an optional base URL option writes each link as an absolute URL.

The feature changes how link targets are written. It does not convert the linked files on its own. A single invocation without `-r` converts one file, as before. To convert a whole tree with one command, the `-r`/`--recursive` option converts every Markdown file under an input directory, and the `-o`/`--outdir` option writes the resulting .docx files under a chosen output directory in the same relative layout as the source tree. Because `--outdir` mirrors the layout, relative links between the converted files stay valid. In recursive mode the Converter knows the whole tree, so it checks link targets against that tree and, with `--link-base-url`, builds absolute links from the root of the output tree.

## Glossary

- **Converter**: The md_to_docx program that reads a Markdown file and writes a .docx file.
- **Source_File**: The Markdown file being converted in the current invocation.
- **Output_File**: The .docx file written by the current invocation.
- **Sibling_Link**: A Markdown link whose target is a relative filesystem path ending in `.md` or `.markdown` (case-insensitive), with or without a `#fragment`. The path may name a file in the same directory as the Source_File or in a subdirectory or parent directory of it.
- **Link_Target**: The destination part of a Markdown link, for example `docs/setup.md#install` in `[Setup](docs/setup.md#install)`.
- **Target_Path**: The Link_Target without its `#fragment`.
- **Fragment**: The part of a Link_Target after `#`.
- **Docx_Target**: The path written into the .docx hyperlink. It is the Target_Path with its Markdown extension replaced by `.docx`. It is a relative path, or, when a Base_URL is given, it is joined to the Base_URL to form an absolute URL.
- **Base_URL**: The absolute `http` or `https` URL of the folder that corresponds to the directory of the Output_File in single-file mode, or to the Output_Root in recursive mode, supplied with the `--link-base-url` option.
- **SharePoint_Mode**: The behavior enabled by the `--sharepoint` option (alias `--sp`), in which each Absolute_Link opens the linked document in Word for the web instead of downloading it.
- **SharePoint_View_URL**: The URL of a SharePoint document library page, whose path ends in `/Forms/<page>.aspx`, for example `https://host/sites/S/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FS%2FShared%20Documents%2FTEST`. When present, its `id` query parameter holds the server-relative path of the folder being viewed.
- **Input_Root**: The directory given as the input argument in recursive mode.
- **Output_Root**: The directory under which output files are written in recursive mode or with `--outdir`. It is the `--outdir` directory when that option is given, and the Input_Root otherwise.
- **Recursive_Mode**: The mode in which the Converter is invoked with `-r` or `--recursive` and the input is a directory.
- **Single_File_Mode**: Any invocation that is not in Recursive_Mode, in which exactly one Source_File is converted.
- **Tree_Source_File**: A file under the Input_Root, at any depth, whose extension is `.md` or `.markdown` (case-insensitive) and which is not inside a Skipped_Directory.
- **Skipped_Directory**: A directory whose name begins with `.` or equals `node_modules`, `venv`, or `__pycache__`. This includes `.git` and `.venv`.
- **Relative_Path**: The path of a file or directory relative to the Input_Root, using `/` as separator.
- **Absolute_Link**: A hyperlink target formed from the Base_URL and the Docx_Target as specified in Requirement 8.
- **Wiki_Link**: An Obsidian-style link of the form `[[Note]]`, `[[Note|Label]]`, or `[[Note#Heading]]`.
- **Link_Style**: The existing `Hyperlink` character style, configured by the `links` section of the style config.
- **Warning**: A message written to standard error that does not stop the conversion.

## Requirements

### Requirement 1: Rewrite Sibling Links

**User Story:** As a documentation author, I want links to sibling Markdown files to become links to the matching .docx files, so that I can click through a converted tree of documents.

#### Acceptance Criteria

1. WHEN the Source_File contains a Sibling_Link, THE Converter SHALL write exactly one hyperlink in the Output_File for that Sibling_Link, with the Docx_Target as its target.
2. WHEN the Target_Path ends in `.md` or `.markdown` in any letter case, THE Converter SHALL replace that extension with `.docx` (lowercase) in the Docx_Target and SHALL leave all other characters of the Target_Path unchanged.
3. WHEN the Link_Target contains a Fragment after the extension, THE Converter SHALL replace the extension that precedes the Fragment and SHALL handle the Fragment as specified in Requirement 2.
4. THE Converter SHALL keep the link text of a Sibling_Link, including any inline formatting in it, as the visible text of the hyperlink.
5. THE Converter SHALL apply the Link_Style to each hyperlink created from a Sibling_Link.
6. WHERE the `--link-base-url` option is not given, THE Converter SHALL write the Docx_Target as a relative path and SHALL NOT convert it to an absolute path.
7. WHEN the Target_Path contains directory segments such as `../` or `sub/`, THE Converter SHALL preserve those segments, in their original order and count, in the Docx_Target.
8. WHEN the Target_Path begins with `./`, THE Converter SHALL write the Docx_Target without the leading `./` and SHALL NOT remove any `./` that appears later in the path.
9. IF the Target_Path does not end in `.md` or `.markdown` (in any letter case, ignoring any Fragment), THEN THE Converter SHALL NOT treat the link as a Sibling_Link and SHALL NOT change its target.
10. WHEN the link text of a Sibling_Link, or the label of a Wiki_Link of the form `[[Note|Label]]`, contains inline formatting (bold, italic, code, or other inline styles), THE Converter SHALL render each formatted run inside the hyperlink with the same formatting that the run has outside a link, plus the Link_Style.

### Requirement 2: Preserve Heading Fragments

**User Story:** As a documentation author, I want links to specific headings in other files to keep their fragment, so that a link can lead to the right section of the target document.

#### Acceptance Criteria

1. WHEN a Sibling_Link contains a non-empty Fragment, THE Converter SHALL define the logical link destination as the Docx_Target, then a single `#` character, then the converted Fragment, with no characters between the Docx_Target filename and the `#`. This criterion defines only the logical destination; the storage of the destination in the hyperlink is specified solely in criterion 6.
2. WHEN a Sibling_Link contains a non-empty Fragment, THE Converter SHALL convert the Fragment to a string that is identical, character for character, to the bookmark name that the Converter generates for the matching heading when it converts the target file.
3. WHEN a Sibling_Link has no Fragment, THE Converter SHALL write a Docx_Target that contains zero `#` characters.
4. IF a Sibling_Link has a `#` character followed by an empty Fragment, THEN THE Converter SHALL write a Docx_Target that contains zero `#` characters.
5. IF a Sibling_Link contains a Fragment that matches no heading in the target file, or the target file cannot be read, THEN THE Converter SHALL still append the converted Fragment to the Docx_Target and SHALL complete the conversion of the Source_File without raising an error.
6. WHEN a Sibling_Link or Wiki_Link points to another document, THE Converter SHALL write the hyperlink as a `w:hyperlink` element whose external relationship target is the Docx_Target file path or URL without a Fragment, and whose `w:anchor` attribute carries the converted Fragment as the bookmark name, and SHALL NOT place `file.docx#bookmark` in the relationship target.
7. WHERE link rewriting is enabled (the `--no-sibling-links` option is not given), THE Converter SHALL generate every heading bookmark name, and therefore every converted Fragment, with at most 40 characters, starting with a letter or underscore, and containing only letters, digits, and underscores.
8. WHERE link rewriting is enabled, WHEN the slug of a heading would exceed 40 characters, THE Converter SHALL truncate the slug and append a short hash suffix computed deterministically from the full slug, so that the bookmark name is at most 40 characters, stays unique among the headings of the document, and is identical on every conversion of the same heading.
9. WHERE the `--no-sibling-links` option is given, THE Converter SHALL generate heading bookmark names exactly as before this feature, using the legacy form `_toc_{idx}_{slug}`, and SHALL NOT apply criteria 7 and 8.

### Requirement 3: Leave Other Links Unchanged

**User Story:** As a documentation author, I want only links to local Markdown files to be rewritten, so that other links keep working as they do today.

#### Acceptance Criteria

1. WHEN a Link_Target begins with a URL scheme (a letter followed by letters, digits, `+`, `-`, or `.`, then `:`, such as `http:`, `https:`, or `mailto:`), THE Converter SHALL write the Link_Target unchanged, including when it ends in `.md` or `.markdown`.
2. WHEN a Link_Target begins with `#`, THE Converter SHALL resolve it as an internal link to the heading in the Output_File whose anchor matches the text after `#`, and SHALL NOT rewrite it.
3. WHEN a Link_Target is a relative path whose extension, compared case-insensitively, is not `.md` or `.markdown`, THE Converter SHALL write the Link_Target unchanged.
4. WHEN a Link_Target is an absolute filesystem path (beginning with `/` or a drive letter such as `C:\`) whose extension, compared case-insensitively, is `.md` or `.markdown`, THE Converter SHALL write the Link_Target unchanged.
5. WHEN a Link_Target is empty or contains only whitespace, THE Converter SHALL write the Link_Target unchanged and SHALL complete the conversion without raising an error.
6. IF a Link_Target is left unchanged under criteria 1 to 5, THEN THE Converter SHALL produce the same link text and destination in the Output_File as it did before this feature was added.

### Requirement 4: Support Wiki Links

**User Story:** As a documentation author who uses Obsidian, I want wiki links between notes to become links to the matching .docx files, so that vault-style Markdown converts the same way as standard links.

#### Acceptance Criteria

1. WHEN the Source_File contains a Wiki_Link of the form `[[Note]]`, THE Converter SHALL write a hyperlink whose Docx_Target is `Note.docx` and whose visible text is `Note`.
2. WHEN the Source_File contains a Wiki_Link of the form `[[Note|Label]]`, THE Converter SHALL write a hyperlink whose Docx_Target is `Note.docx` and whose visible text is `Label`.
3. WHEN the Source_File contains a Wiki_Link of the form `[[Note#Heading]]`, THE Converter SHALL write a hyperlink whose Docx_Target is `Note.docx` (without a Fragment) and whose `w:anchor` attribute carries the bookmark name that the Converter generates for `Heading`, as specified in Requirement 2, criterion 6, and whose visible text is `Note#Heading`.
4. WHEN the Source_File contains a Wiki_Link of the form `[[#Heading]]`, THE Converter SHALL write an internal hyperlink to the bookmark that the Converter generates for `Heading` in the Output_File, with visible text `Heading`.
5. WHEN the Wiki_Link note name has a file extension of `.md` or `.markdown`, THE Converter SHALL replace that extension with `.docx` in the Docx_Target.
6. WHEN the Wiki_Link note name contains a file extension other than `.md` or `.markdown`, THE Converter SHALL append `.docx` to the full note name in the Docx_Target.
7. IF a Wiki_Link is malformed, meaning it has an empty note name and no heading (`[[]]`), an empty label after `|` (`[[Note|]]`), or no closing `]]`, THEN THE Converter SHALL leave the original text unchanged as plain text in the Output_File, SHALL NOT write a hyperlink for it, and SHALL continue converting the remainder of the Source_File.
8. IF a Wiki_Link of the form `[[#Heading]]` names a heading that does not exist in the Source_File, THEN THE Converter SHALL write the visible text `Heading` as plain text without a hyperlink and SHALL continue converting the remainder of the Source_File.

### Requirement 5: Handle Special Characters in Paths

**User Story:** As a documentation author, I want link targets with spaces or special characters to work, so that file names such as `My Notes.md` do not produce broken links.

#### Acceptance Criteria

1. WHEN a Target_Path contains percent-encoded sequences (for example `%20`), THE Converter SHALL decode each valid sequence (a `%` followed by two hexadecimal digits, interpreted as UTF-8) before replacing the extension, and SHALL percent-encode the resulting Docx_Target for use as the hyperlink address so that the address contains no unencoded spaces or characters outside the URL-safe set (letters, digits, `-`, `.`, `_`, `~`, and the path separator `/`).
2. WHEN a Target_Path is wrapped in angle brackets, such as `[Notes](<My Notes.md>)`, THE Converter SHALL treat the text between the opening `<` and the first closing `>` as the Target_Path, excluding the brackets, and SHALL preserve spaces inside it.
3. WHEN a Target_Path contains backslash (`\`) path separators, THE Converter SHALL replace every backslash with a forward slash (`/`) in the Docx_Target.
4. IF a Target_Path contains a `%` not followed by two hexadecimal digits, or a percent-encoded sequence that is not valid UTF-8, THEN THE Converter SHALL leave that sequence undecoded, SHALL continue converting the document, and SHALL produce a hyperlink for that link without raising an error.
5. IF a Target_Path wrapped in angle brackets has no closing `>`, THEN THE Converter SHALL leave the link text unchanged as plain text, SHALL continue converting the remainder of the document, and SHALL NOT produce a hyperlink for that link.

### Requirement 6: Warn About Missing Targets

**User Story:** As a documentation author, I want a warning when a linked Markdown file does not exist, so that I can find broken links before I distribute the documents.

#### Acceptance Criteria

1. WHEN a Sibling_Link resolves, relative to the directory of the Source_File, to a file that does not exist, THE Converter SHALL write exactly one Warning for that link to the standard error stream, and the Warning SHALL contain the Source_File path and the Target_Path as written in the Source_File.
2. WHEN a Sibling_Link resolves to a file that does not exist, THE Converter SHALL write the hyperlink with the Docx_Target, and THE Converter SHALL continue converting the remaining content of the Source_File.
3. WHEN a Sibling_Link resolves to a file that exists, THE Converter SHALL write zero Warnings for that link.
4. IF a Sibling_Link resolves to a path outside the directory tree that contains the Source_File, THEN THE Converter SHALL write the hyperlink with the Docx_Target, and THE Converter SHALL determine whether to write a Warning by checking whether the resolved file exists, using the same rule as criteria 1 and 3.
5. WHEN a Sibling_Link whose target does not exist appears N times in the Source_File, THE Converter SHALL write N Warnings, one for each occurrence, and the conversion SHALL finish with the same exit status as a conversion that has no missing targets.

### Requirement 7: Disable Link Rewriting

**User Story:** As a documentation author, I want to turn off link rewriting, so that I can keep Markdown links as written when I need that.

#### Acceptance Criteria

1. WHEN the Converter is invoked with the `--no-sibling-links` command-line option, THE Converter SHALL accept the option as a boolean flag that takes no value and SHALL complete the conversion without an argument-parsing error.
2. WHERE the `--no-sibling-links` option is given, THE Converter SHALL produce exactly the same Output_File content as the Converter produced before this feature existed, so that Sibling_Links and Wiki_Links are rendered as plain text and not as hyperlinks, and all other links are rendered as they were before this feature.
3. WHERE the `--no-sibling-links` option is not given, THE Converter SHALL rewrite Sibling_Links and Wiki_Links as specified in Requirements 1 to 5.
4. WHEN the Converter is invoked with `--help`, THE Converter SHALL include the `--no-sibling-links` option name and a one-line description of its effect in the help output.
5. IF the `--no-sibling-links` option is given with a value (for example `--no-sibling-links=yes`), THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message indicating that the option takes no value to standard error, and SHALL NOT create or modify the output file.

### Requirement 8: Absolute Links with a Base URL

**User Story:** As a documentation author whose readers view the .docx files in a browser from SharePoint, I want links written as absolute URLs, so that the links work where relative paths cannot be resolved.

#### Acceptance Criteria

1. THE Converter SHALL accept an optional `--link-base-url` command-line option that takes exactly one value, the Base_URL.
2. WHEN the Converter is invoked with `--help`, THE Converter SHALL include the `--link-base-url` option name and a one-line description of its effect in the help output.
3. WHERE the `--link-base-url` option is not given, THE Converter SHALL write relative Docx_Targets exactly as specified in Requirements 1 to 5.
4. WHERE the `--link-base-url` option is given, THE Converter SHALL write the hyperlink target of each Sibling_Link and Wiki_Link as an Absolute_Link, formed by joining the Base_URL, with exactly one `/` between the two parts, and the percent-encoded Docx_Target resolved relative to the Base_URL folder. In Single_File_Mode the Base_URL folder corresponds to the directory of the Output_File. In Recursive_Mode the Base_URL folder corresponds to the Output_Root, as specified in Requirement 12.
5. WHERE the `--link-base-url` option is given, THE Converter SHALL remove `./` segments and resolve each `../` segment against the preceding path segment when forming the Absolute_Link.
6. WHERE the `--link-base-url` option is given and the Sibling_Link or Wiki_Link has a non-empty Fragment, THE Converter SHALL carry the Fragment, as specified in Requirement 2, in the `w:anchor` attribute of the hyperlink, with the Absolute_Link without Fragment as the external relationship target, so that the visible URL semantics are the same as `<Absolute_Link>#<Fragment>`.
7. WHEN the Source_File links to `b/c.md` and the Base_URL is `https://x/docs/`, THE Converter SHALL write the Absolute_Link `https://x/docs/b/c.docx`.
8. WHEN the Base_URL is given with a trailing `/` and when it is given without one, THE Converter SHALL write the same Absolute_Link for the same Link_Target.
9. IF resolving `../` segments in a Docx_Target would climb above the root path of the Base_URL, THEN THE Converter SHALL still write an Absolute_Link, formed by discarding each `../` segment that has no preceding path segment to resolve against, and SHALL write one Warning for that link to the standard error stream containing the Source_File path and the Target_Path as written in the Source_File.
10. WHEN the Base_URL contains percent-encoded sequences, THE Converter SHALL use those sequences unchanged and SHALL NOT percent-encode the `%` character of a valid sequence again.
11. WHEN the Base_URL contains a space or another character outside the URL-safe set that is not part of a valid percent-encoded sequence, THE Converter SHALL percent-encode that character once in the Absolute_Link.
12. IF the Base_URL is not an absolute URL with the scheme `http` or `https` (compared case-insensitively) and a non-empty host, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message naming the invalid Base_URL to the standard error stream, and SHALL NOT create or modify the Output_File.
13. WHERE both the `--link-base-url` option and the `--no-sibling-links` option are given, THE Converter SHALL apply the `--no-sibling-links` behavior specified in Requirement 7, producing the same output as before this feature existed, and SHALL NOT form any Absolute_Link.
14. WHERE the `--link-base-url` option is given, THE Converter SHALL write each Link_Target that is not a Sibling_Link or Wiki_Link, including external URLs, `#` anchors, and links to non-Markdown files, with the same destination as it writes when the option is not given.

### Requirement 9: Preserve Directory Layout

**User Story:** As a documentation author, I want the output tree to match the source tree, so that the relative links work after conversion.

#### Acceptance Criteria

1. THE Converter SHALL write the Docx_Target with the same relative directory segments as the Target_Path, so that the path from the Output_File to the linked .docx file mirrors the path from the Source_File to the linked Markdown file.
2. WHEN a Source_File at `docs/a.md` links to `b/c.md` and the Output_File is written to `out/docs/a.docx`, THE Converter SHALL write the Docx_Target `b/c.docx`, which resolves to `out/docs/b/c.docx`.
3. WHEN the Output_File is written to a directory that differs from the Source_File directory, THE Converter SHALL write the same Docx_Target that it writes when the Output_File and the Source_File are in the same directory.
4. THE Converter SHALL NOT create, read, or modify the file that a Docx_Target names while converting a Source_File in Single_File_Mode.
5. WHERE the `--outdir` option is given in Recursive_Mode, THE Converter SHALL write each Output_File at the same Relative_Path under the Output_Root as its Tree_Source_File has under the Input_Root, so that every Docx_Target resolves to the Output_File of the linked Tree_Source_File.

### Requirement 10: Convert a Directory Tree

**User Story:** As a documentation author, I want to convert a whole directory of Markdown files with one command, so that I do not have to invoke the Converter once per file.

#### Acceptance Criteria

1. THE Converter SHALL accept an optional `-r` / `--recursive` command-line option that is a boolean flag taking no value.
2. WHEN the Converter is invoked with `-r` and an input path that is a directory, THE Converter SHALL convert every Tree_Source_File, at any depth, to an Output_File with the same file stem and the extension `.docx`.
3. WHEN the Converter is invoked with `-r` and an input path that is a directory, THE Converter SHALL treat the input path as the Input_Root and SHALL NOT require a positional output argument.
4. IF the Converter is invoked with `-r`, an input directory, and a positional output argument, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message stating that the output argument cannot be combined with `-r` to standard error, and SHALL NOT create or modify any file.
5. IF the input path is a directory and the `-r` option is not given, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message that names the directory and suggests the `-r` option, and SHALL NOT create or modify any file.
6. WHEN the Converter is invoked with `-r` and an input path that is a file, THE Converter SHALL behave as Single_File_Mode, including the requirement for a positional output argument unless `--outdir` is given.
7. IF the input path does not exist, THEN THE Converter SHALL exit with a non-zero exit status and SHALL write an error message naming the path, with or without `-r`.
8. THE Converter SHALL match the extensions `.md` and `.markdown` case-insensitively when it selects Tree_Source_Files.
9. THE Converter SHALL exclude from the conversion every file and directory whose name begins with `.`, and every directory named `node_modules`, `venv`, or `__pycache__`, together with all of their contents.
10. THE Converter SHALL NOT follow symbolic links that point to directories while it searches the Input_Root, and SHALL convert a symbolic link that points to a file only when the link itself is a Tree_Source_File.
11. THE Converter SHALL convert the Tree_Source_Files in ascending order of their Relative_Path, compared as Unicode code points.
12. WHEN the Converter finishes converting a Tree_Source_File, THE Converter SHALL write exactly one progress line to standard output that contains the Relative_Path of the Tree_Source_File and the path of its Output_File.
13. IF the Input_Root contains no Tree_Source_File, THEN THE Converter SHALL write a Warning that names the Input_Root, SHALL create no Output_File, and SHALL exit with exit status 0.
14. IF the conversion of one Tree_Source_File fails, THEN THE Converter SHALL write an error message containing the Relative_Path and the cause to standard error, SHALL continue converting the remaining Tree_Source_Files, and SHALL NOT write or leave a partial Output_File for the failed Tree_Source_File.
15. IF the conversion of at least one Tree_Source_File fails, THEN THE Converter SHALL exit with a non-zero exit status after it has attempted every Tree_Source_File.
16. WHEN every Tree_Source_File is converted successfully, THE Converter SHALL exit with exit status 0.
17. WHEN the Converter finishes in Recursive_Mode, THE Converter SHALL write a summary line to standard output that states the number of files converted and the number of files that failed.
18. WHEN the Converter is invoked with `--help`, THE Converter SHALL include the `-r` and `--recursive` option names and a one-line description of their effect in the help output.
19. THE Converter SHALL apply the `--style`, `--fetch-remote-images`, `--no-sibling-links`, and `--link-base-url` options to every Tree_Source_File in Recursive_Mode.
20. WHEN the Converter is invoked with `--dump-config`, THE Converter SHALL behave as before and SHALL ignore `-r` and `--outdir`.

### Requirement 11: Output Directory

**User Story:** As a documentation author, I want to choose where the converted files are written, so that the .docx files do not mix with my Markdown sources.

#### Acceptance Criteria

1. THE Converter SHALL accept an optional `-o`/`--outdir` command-line option that takes exactly one value, a directory path.
2. WHEN the Converter is invoked with `--outdir DIR` in Recursive_Mode, THE Converter SHALL write the Output_File of each Tree_Source_File to `DIR/<Relative_Path with its extension replaced by .docx>`, so that the Input_Root maps to DIR.
3. WHEN the Converter is invoked with `--outdir DIR` and an input path that is a file, THE Converter SHALL write the Output_File to `DIR/<stem>.docx` and SHALL NOT require a positional output argument.
4. WHERE `--outdir` is not given in Recursive_Mode, THE Converter SHALL write each Output_File in the same directory as its Tree_Source_File, with the same stem and the extension `.docx`.
5. WHEN a directory required for an Output_File does not exist, THE Converter SHALL create it, including missing parent directories, before it writes the Output_File.
6. IF the Converter is invoked with `--outdir` and a positional output argument, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message stating that the two cannot be combined to standard error, and SHALL NOT create or modify any file.
7. WHEN an Output_File already exists, THE Converter SHALL overwrite it, subject to the locked-file check that Single_File_Mode performs.
8. IF the `--outdir` directory cannot be created or is not writable, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message naming the directory to standard error, and SHALL NOT convert any file.
9. IF the `--outdir` path exists and is not a directory, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message naming the path to standard error, and SHALL NOT convert any file.
10. WHERE the `--outdir` directory is inside the Input_Root, THE Converter SHALL select Tree_Source_Files only by the rules of Requirement 10, so that the generated .docx files are never selected for conversion.
11. WHEN two Tree_Source_Files in the same directory have the same stem (for example `a.md` and `a.markdown`), THE Converter SHALL write an error message naming both Relative_Paths, SHALL NOT convert either file, SHALL continue with the remaining Tree_Source_Files, and SHALL exit with a non-zero exit status.
12. WHEN the Converter is invoked with `--help`, THE Converter SHALL include the `-o` and `--outdir` option names and a one-line description of their effect in the help output.
13. WHEN the Converter is invoked with `-o DIR`, THE Converter SHALL behave identically to an invocation with `--outdir DIR`, and THE Converter SHALL list both `-o` and `--outdir` in the help output.

### Requirement 12: Link Rewriting in Recursive Mode

**User Story:** As a documentation author converting a whole tree, I want link checks and absolute links to be based on the tree, so that every link points to the converted file of its target.

#### Acceptance Criteria

1. WHILE the Converter is in Recursive_Mode, THE Converter SHALL decide whether a Sibling_Link target exists, for the Warning of Requirement 6, by checking whether the resolved file exists on the filesystem, using the Input_Root as the reference for the tree.
2. WHILE the Converter is in Recursive_Mode, WHEN a Sibling_Link resolves to a Markdown file that exists but is not a Tree_Source_File (because it is in a Skipped_Directory or outside the Input_Root), THE Converter SHALL write one Warning for that link that contains the Source_File path and the Target_Path as written, and SHALL still write the hyperlink.
3. WHILE the Converter is in Recursive_Mode, WHERE the `--link-base-url` option is given, THE Converter SHALL treat the Base_URL as the URL of the Output_Root.
4. WHILE the Converter is in Recursive_Mode, WHERE the `--link-base-url` option is given, THE Converter SHALL form the Absolute_Link of each Sibling_Link by resolving the Target_Path against the Relative_Path of the Source_File, replacing the Markdown extension with `.docx`, and joining the result to the Base_URL as specified in Requirement 8.
5. WHEN the Source_File at Relative_Path `a/b/x.md` links to `../y.md` and the Base_URL is `https://x/docs/`, THE Converter SHALL write the Absolute_Link `https://x/docs/a/y.docx`.
6. WHILE the Converter is in Recursive_Mode, IF resolving a Target_Path against the Relative_Path of the Source_File climbs above the Input_Root, THEN THE Converter SHALL write the Absolute_Link and the Warning as specified in Requirement 8, criterion 9, treating the Output_Root as the root of the Base_URL folder.
7. WHILE the Converter is in Recursive_Mode, WHERE the `--link-base-url` option is not given, THE Converter SHALL write relative Docx_Targets as specified in Requirements 1 to 5, and the links SHALL resolve to the Output_File of the linked Tree_Source_File because the Output_Root preserves the Relative_Path layout of the Input_Root.
8. WHILE the Converter is in Single_File_Mode, THE Converter SHALL define the Base_URL folder as the directory of the Output_File, as specified in Requirement 8.
9. WHILE the Converter is in Recursive_Mode, THE Converter SHALL resolve each Fragment as specified in Requirement 2.

### Requirement 13: Documentation and Testing

**User Story:** As a maintainer, I want the feature documented and tested, so that the behavior stays stable.

#### Acceptance Criteria

1. THE README SHALL describe the link rewriting rules, the `--no-sibling-links` option including its default behavior when the option is omitted, the requirement that .docx files keep the same relative layout as the Markdown files, and the `--link-base-url` option, including an example that uses a SharePoint folder URL and a statement that relative links work only with local or synced folders.
2. THE CHANGELOG SHALL contain one entry under the unreleased section that names the sibling link rewriting feature and the `--no-sibling-links` option.
3. THE test suite SHALL contain at least one automated test for each of Requirements 1 to 12 that covers the nominal path and, where the requirement defines an error or exclusion case, at least one test for that case.
4. THE README SHALL document the `-r` / `--recursive` option and the `-o`/`--outdir` option, including the skipped directories, the mirrored output layout, the exit status on partial failure, and a SharePoint workflow example that uses the command `md-to-docx -r docs/ --outdir out/ --link-base-url https://.../docs/`.
5. THE CHANGELOG SHALL contain an entry under the unreleased section that names the `-r` / `--recursive` option and the `--outdir` option.
6. THE test suite SHALL contain automated tests for Requirements 10 to 12 that cover nested directories, each Skipped_Directory name, a directory symbolic link that would form a cycle, case-insensitive extensions, sorted processing order, an input directory without `-r`, an empty tree, partial failure with a non-zero exit status while the other files are still converted, `--outdir` mirroring including creation of missing directories, `--outdir` combined with a positional output, `--outdir` inside the Input_Root, `--outdir` with a single file input, and the Base_URL taken from the Output_Root (`a/b/x.md` linking `../y.md` gives `<base>/a/y.docx`).
7. THE test suite SHALL verify that, FOR ALL trees of Markdown files generated with nested directories, mixed-case extensions, and Skipped_Directories, the set of Relative_Paths of the Output_Files under the Output_Root is exactly the set of Relative_Paths of the Tree_Source_Files with each extension replaced by `.docx`, with a one-to-one mapping and the same stem for every file (mirrored-set property).
8. WHEN the test suite is run, THE test suite SHALL pass with 0 failures and 0 errors.
9. THE test suite SHALL verify that, FOR ALL Target_Path values that end in lowercase `.md`, converting the Target_Path to a Docx_Target and then replacing the trailing `.docx` with `.md` produces the original Target_Path (round-trip property).
10. THE test suite SHALL verify that, FOR ALL Target_Path values, converting the Target_Path to a Docx_Target twice produces the same result as converting it once (idempotence property).
11. WHERE a property-based test is run, THE test suite SHALL execute each property from criteria 9 and 10 on at least 100 generated Target_Path values, including values with an empty fragment, a fragment, nested directories, and uppercase extensions.
12. THE test suite SHALL verify that, FOR ALL Target_Path values and FOR ALL valid Base_URL values, the Absolute_Link written with the Base_URL with a trailing `/` is identical to the Absolute_Link written with the Base_URL without a trailing `/` (trailing-slash consistency property).
13. THE test suite SHALL verify that, FOR ALL Sibling_Link values that do not climb above the Base_URL root, the Absolute_Link equals the percent-encoded, resolved Docx_Target joined to the Base_URL, and contains no `#` fragment (Base_URL round-trip property).
14. THE test suite SHALL verify that, FOR ALL Base_URL values, joining an already percent-encoded Base_URL a second time produces the same Absolute_Link as joining it once (no double-encoding property).
15. THE test suite SHALL verify that a Sibling_Link label and a Wiki_Link label containing bold, italic, or code runs are rendered inside the hyperlink with the same formatting as outside a link, plus the Link_Style (Requirement 1, criterion 10).
16. THE test suite SHALL verify that a hyperlink to another document has the Docx_Target without Fragment as its external relationship target and the converted Fragment in `w:anchor`, and that no relationship target contains `#` (Requirement 2, criterion 6).
17. THE test suite SHALL verify that, FOR ALL headings, each bookmark name generated with link rewriting enabled has at most 40 characters, starts with a letter or underscore, and contains only letters, digits, and underscores (Requirement 2, criterion 7).
18. THE test suite SHALL verify that, FOR ALL headings whose slug exceeds 40 characters, the truncated bookmark name is identical across repeated conversions and unique among the headings of one document (Requirement 2, criterion 8).
19. THE test suite SHALL verify that, WHERE the `--no-sibling-links` option is given, the Output_File content is identical to the content produced before this feature for the same Source_File, including legacy `_toc_{idx}_{slug}` heading bookmark names (Requirement 7, criterion 2, and Requirement 2, criterion 9).

### Requirement 14: SharePoint Links

**User Story:** As a documentation author who publishes the .docx files to SharePoint, I want the links to open the linked document in Word for the web, and I want to paste the folder address shown in my browser, so that readers can click through the documents without downloading them.

#### Acceptance Criteria

1. THE Converter SHALL accept an optional `--sharepoint` command-line option, with the alias `--sp`, that takes no value, and SHALL list both names with a one-line description in the help output.
2. IF the `--sharepoint` option is given with a value (for example `--sharepoint=yes`), THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message to standard error, and SHALL NOT create or modify any file.
3. IF the `--sharepoint` option is given without the `--link-base-url` option, THEN THE Converter SHALL exit with a non-zero exit status, SHALL write an error message stating that `--sharepoint` requires `--link-base-url` to standard error, and SHALL NOT create or modify any file.
4. WHERE the `--sharepoint` option and the `--link-base-url` option are both given, THE Converter SHALL append `?web=1` to every Absolute_Link written for a Sibling_Link or Wiki_Link, so that the external relationship target ends in `.docx?web=1`.
5. WHERE SharePoint_Mode is active, THE Converter SHALL carry the converted Fragment in the `w:anchor` attribute as specified in Requirement 2, criterion 6, and SHALL NOT place `#` in the relationship target.
6. WHERE SharePoint_Mode is active and the `--link-base-url` value is a SharePoint_View_URL with a non-empty `id` query parameter, THE Converter SHALL use as Base_URL the scheme and host of the value followed by the percent-decoded `id` path, percent-encoded once as specified in Requirement 8, criteria 10 and 11, and SHALL NOT write the Warning about a dropped query.
7. WHERE SharePoint_Mode is active and the `--link-base-url` value is a SharePoint_View_URL without an `id` query parameter, THE Converter SHALL use as Base_URL the value with the `/Forms/<page>.aspx` suffix and the query removed, which is the root folder of the document library.
8. WHERE SharePoint_Mode is not active and the `--link-base-url` value is a SharePoint_View_URL, THE Converter SHALL keep the behavior of Requirement 8 and SHALL write one Warning that suggests the `--sharepoint` option.
9. WHERE SharePoint_Mode is active, THE Converter SHALL write each Link_Target that is not a Sibling_Link or Wiki_Link with the same destination as it writes when the option is not given.
10. WHERE the `--sharepoint` option and the `--no-sibling-links` option are both given, THE Converter SHALL apply the `--no-sibling-links` behavior specified in Requirement 7 and SHALL NOT form any Absolute_Link.
11. THE README SHALL document the `--sharepoint` and `--sp` options, including an example that passes a SharePoint_View_URL, and a statement that Word for the web may open a linked document without moving to the linked heading.
12. THE CHANGELOG SHALL contain an entry under the unreleased section that names the `--sharepoint` option.
13. THE test suite SHALL contain automated tests for criteria 1 to 10 and SHALL verify that, FOR ALL Sibling_Links and valid Base_URL values in SharePoint_Mode, the relationship target equals the Absolute_Link written without SharePoint_Mode followed by `?web=1`, contains exactly one `?`, and contains no `#` (SharePoint suffix property).
