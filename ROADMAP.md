# Roadmap

## v0.1 (current)

- [x] Copy slides between PPTX files preserving layouts, masters, and themes
- [x] Handle images, video, audio, and other media
- [x] Media deduplication by content hash
- [x] Preserve speaker notes with notes master
- [x] Globally unique OOXML IDs (sldMasterId, sldLayoutId)
- [x] Correct OOXML element ordering in presentation.xml
- [x] Strip comment references from copied slides
- [x] CLI tool (`pptx-merge`)

## v0.2

- [ ] Chart support (copy chart parts and data)
- [ ] OLE object support
- [ ] Slide reordering after merge
- [ ] Improved error reporting for corrupt source files

## v0.3

- [ ] Text search/replace across slides
- [ ] Delete specific slides from a deck
- [ ] Clone slides within the same deck

## v1.0

- [ ] Stable public API
- [ ] Comprehensive test suite with real-world PPTX fixtures
- [ ] Published to PyPI
