"""Test cho công cụ kiểm hiệu ứng PPTX của bản Việt."""

import contextlib
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tools" / "vi"))

import kiem_hieu_ung  # noqa: E402

NAMESPACES = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
    'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'
)

# Cấu trúc sao theo file do svg_to_pptx.py xuất ra.
FADE = '<p:transition p14:dur="400"><p:fade/></p:transition>'
PUSH = '<p:transition p14:dur="400"><p:push dir="r"/></p:transition>'
MORPH = (
    '<mc:AlternateContent><mc:Choice Requires="p159"><p:transition p14:dur="400">'
    '<p159:morph option="byObject"/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'
)
SOUND_ONLY = '<p:transition><p:sndAc><p:stSnd><p:snd r:embed="rId9" name="a.wav"/></p:stSnd></p:sndAc></p:transition>'


def effect(node_type, spid=5, klass="entr", dur=400):
    return (
        f'<p:par><p:cTn id="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="0" presetID="10" presetClass="{klass}" presetSubtype="0" fill="hold" '
        f'nodeType="{node_type}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="0" dur="{dur}"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
        f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
    )


def main_seq(*rows):
    return f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(rows)}</p:childTnLst></p:cTn>'


def interactive_seq(trigger_spid=8, target_spid=11):
    return (
        f'<p:cTn id="3" restart="whenNotActive" fill="hold" nodeType="interactiveSeq">'
        f'<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_spid}"/></p:tgtEl>'
        f'</p:cond></p:stCondLst><p:childTnLst>{effect("clickEffect", target_spid)}</p:childTnLst></p:cTn>'
    )


def timing(*sequences):
    seqs = "".join(f'<p:seq concurrent="1" nextAc="seek">{seq}</p:seq>' for seq in sequences)
    return (
        f'<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst>{seqs}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


AUDIO = (
    # Nút phát thuyết minh do narration.py chèn cạnh mainSeq: không phải hiệu ứng đối tượng.
    '<p:audio><p:cMediaNode vol="80000"><p:cTn id="16" fill="hold" display="0"><p:stCondLst>'
    '<p:cond delay="400"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="12"/></p:tgtEl>'
    '</p:cMediaNode></p:audio>'
)


def narration_timing(*sequences):
    seqs = "".join(f'<p:seq concurrent="1" nextAc="seek">{seq}</p:seq>' for seq in sequences)
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst>{seqs}{AUDIO}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


def slide(transition=FADE, timing_xml="", lines=(), hidden=False):
    paragraphs = "".join(f"<a:p><a:r><a:t>{line}</a:t></a:r></a:p>" for line in lines)
    body = f'<p:sp><p:txBody>{paragraphs}</p:txBody></p:sp>' if paragraphs else ""
    show = ' show="0"' if hidden else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><p:sld {NAMESPACES}{show}><p:cSld><p:spTree>{body}'
        f'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{transition}{timing_xml}</p:sld>'
    )


def reveal_slide(clicks=3, node_type="clickEffect"):
    return slide(timing_xml=timing(main_seq(*[effect(node_type, 5 + i) for i in range(clicks)])))


def build_deck(folder: Path, slides, name="bai.pptx") -> Path:
    path = folder / name
    ids = "".join(f'<p:sldId id="{256 + i}" r:id="rId{i + 1}"/>' for i in range(len(slides)))
    rels = "".join(
        f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/slide" Target="slides/slide{i + 1}.xml"/>' for i in range(len(slides))
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/presentation.xml",
                         f'<?xml version="1.0"?><p:presentation {NAMESPACES}><p:sldIdLst>{ids}</p:sldIdLst>'
                         f'</p:presentation>')
        archive.writestr("ppt/_rels/presentation.xml.rels",
                         '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/'
                         f'package/2006/relationships">{rels}</Relationships>')
        # Ghi slide theo thứ tự ngược để chắc công cụ đọc thứ tự từ presentation.xml.
        for index in reversed(range(len(slides))):
            archive.writestr(f"ppt/slides/slide{index + 1}.xml", slides[index])
    return path


def plain_deck(count=6):
    return [slide() for _ in range(count)]


class CheckerCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def run_tool(self, *argv):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = kiem_hieu_ung.main([str(arg) for arg in argv])
        lines = out.getvalue().splitlines()
        self.assertEqual(len(lines), 1, out.getvalue())
        return code, json.loads(lines[0])

    def check(self, slides, *flags):
        return self.run_tool(build_deck(self.folder, slides), *flags)


class LevelKhongTest(CheckerCase):
    def test_fade_only_deck_passes(self):
        code, data = self.check(plain_deck(), "--muc", "khong")
        self.assertEqual(code, 0)
        self.assertTrue(data["ready"])
        self.assertIsNone(data["error"])
        self.assertEqual(data["transitions"], {"fade": 6})

    def test_sound_only_transition_counts_as_plain(self):
        deck = plain_deck()
        deck[2] = slide(transition=SOUND_ONLY)
        code, data = self.check(deck, "--muc", "khong")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["transitions"].get("none"), 1)

    def test_object_effect_fails(self):
        deck = plain_deck()
        deck[1] = reveal_slide(1)
        code, data = self.check(deck, "--muc", "khong")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "muc")
        self.assertIn("trang 2", data["error"]["message"])

    def test_non_fade_transition_fails(self):
        deck = plain_deck()
        deck[3] = slide(transition=PUSH)
        code, data = self.check(deck, "--muc", "khong")
        self.assertEqual(code, 1)
        self.assertIn("trang 4", data["error"]["message"])


class LevelVuaTest(CheckerCase):
    def test_fade_only_deck_fails_with_counts(self):
        code, data = self.check(plain_deck(), "--muc", "vua")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "muc")
        self.assertIn("Mức vừa", data["error"]["message"])
        self.assertEqual((data["content_slides"], data["slides_with_effects"]), (4, 0))

    def test_threshold_is_thirty_percent_of_content_slides(self):
        deck = plain_deck(12)  # 10 trang nội dung, cần 3
        deck[1] = reveal_slide()
        deck[2] = reveal_slide()
        code, _ = self.check(deck, "--muc", "vua")
        self.assertEqual(code, 1)
        deck[3] = reveal_slide()
        code, data = self.check(deck, "--muc", "vua")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["slides_with_effects"], 3)

    def test_cover_and_ending_effects_do_not_count(self):
        deck = plain_deck()
        deck[0] = reveal_slide()
        deck[-1] = reveal_slide()
        _, data = self.check(deck, "--muc", "vua")
        self.assertEqual(data["slides_with_effects"], 0)

    def test_too_many_clicks_on_one_slide_warns(self):
        deck = plain_deck()
        deck[1] = reveal_slide(2)
        deck[2] = reveal_slide(9)
        code, data = self.check(deck, "--muc", "vua")
        self.assertEqual(code, 0, data)
        self.assertEqual((data["max_click_steps"], data["max_click_slide"]), (9, 3))
        self.assertTrue(any("9 bước bấm" in w for w in data["warnings"]))

    def test_long_deck_without_highlight_transition_warns(self):
        deck = plain_deck(8)
        for index in (1, 2, 3):
            deck[index] = reveal_slide()
        code, data = self.check(deck, "--muc", "vua")
        self.assertEqual(code, 0, data)
        self.assertTrue(any("chuyển trang nổi bật" in w for w in data["warnings"]))
        deck[4] = slide(transition=PUSH)
        _, data = self.check(deck, "--muc", "vua")
        self.assertFalse(any("chuyển trang nổi bật" in w for w in data["warnings"]))

    def test_narration_media_nodes_are_not_effects(self):
        deck = plain_deck()
        for index in range(1, 5):
            deck[index] = slide(timing_xml=narration_timing())
        _, data = self.check(deck, "--muc", "vua")
        self.assertEqual((data["slides_with_effects"], data["max_click_steps"]), (0, 0))

    def test_narration_next_to_main_sequence_keeps_effect_count(self):
        deck = plain_deck()
        deck[1] = slide(timing_xml=narration_timing(main_seq(effect("afterEffect", 5), effect("afterEffect", 6))))
        deck[2] = slide(timing_xml=narration_timing(main_seq(effect("withEffect", 5))))
        code, data = self.check(deck, "--muc", "vua")
        self.assertEqual(code, 0, data)
        self.assertEqual((data["slides_with_effects"], data["max_click_steps"]), (2, 0))

    def test_exit_only_slide_counts_as_effect(self):
        deck = plain_deck()
        deck[1] = slide(timing_xml=timing(main_seq(effect("clickEffect", klass="exit"))))
        deck[2] = slide(timing_xml=timing(main_seq(effect("clickEffect", klass="emph"))))
        _, data = self.check(deck, "--muc", "vua")
        self.assertEqual(data["slides_with_effects"], 2)

    def test_interactive_sequence_before_main_sequence(self):
        deck = plain_deck()
        deck[1] = slide(timing_xml=timing(interactive_seq(), main_seq(effect("clickEffect"), effect("clickEffect"))))
        deck[2] = reveal_slide(1)
        _, data = self.check(deck, "--muc", "vua")
        self.assertEqual((data["max_click_steps"], data["max_click_slide"], data["flip_cards"]), (2, 2, 1))

    def test_timing_inside_alternate_content_is_read_once(self):
        choice = timing(main_seq(effect("clickEffect"), effect("clickEffect")))
        wrapped = (f'<mc:AlternateContent><mc:Choice Requires="p14">{choice}</mc:Choice>'
                   f'<mc:Fallback>{choice}</mc:Fallback></mc:AlternateContent>')
        deck = plain_deck()
        deck[1] = slide(timing_xml=wrapped)
        deck[2] = reveal_slide(1)
        _, data = self.check(deck, "--muc", "vua")
        self.assertEqual((data["slides_with_effects"], data["max_click_steps"]), (2, 2))

    def test_hidden_slides_are_skipped(self):
        deck = plain_deck()
        deck.insert(2, slide(transition=PUSH, hidden=True))
        code, data = self.check(deck, "--muc", "khong")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["slides"], 6)


class LevelNhieuTest(CheckerCase):
    def rich_deck(self):
        deck = plain_deck()
        deck[1] = reveal_slide()
        deck[2] = slide(timing_xml=timing(interactive_seq()))
        deck[3] = slide(transition=PUSH)
        deck[4] = slide(transition=MORPH)
        return deck

    def test_rich_deck_passes_and_reads_morph_inside_alternate_content(self):
        code, data = self.check(self.rich_deck(), "--muc", "nhieu")
        self.assertEqual(code, 0, data)
        self.assertEqual((data["morph_slides"], data["flip_cards"]), (1, 1))
        self.assertEqual(data["transitions"], {"fade": 4, "push": 1, "morph": 1})
        self.assertEqual(data["slides_with_effects"], 3)
        self.assertEqual(data["warnings"], [])

    def test_missing_morph_only_warns(self):
        deck = self.rich_deck()
        deck[4] = reveal_slide()
        code, data = self.check(deck, "--muc", "nhieu")
        self.assertEqual(code, 0, data)
        self.assertTrue(any("Morph" in w for w in data["warnings"]))

    def test_below_half_fails(self):
        deck = plain_deck(12)  # 10 trang nội dung, cần 5
        deck[1] = reveal_slide()
        deck[2] = slide(transition=MORPH)
        code, data = self.check(deck, "--muc", "nhieu")
        self.assertEqual(code, 1)
        self.assertIn("cần ít nhất 5", data["error"]["message"])

    def test_flip_card_clicks_are_not_main_sequence_steps(self):
        _, data = self.check(self.rich_deck(), "--muc", "nhieu")
        self.assertEqual((data["max_click_steps"], data["max_click_slide"]), (3, 2))

    def test_no_flip_card_warns(self):
        deck = self.rich_deck()
        deck[2] = reveal_slide()
        code, data = self.check(deck, "--muc", "nhieu")
        self.assertEqual(code, 0, data)
        self.assertTrue(any("hiện đáp án" in w for w in data["warnings"]))


class VideoTest(CheckerCase):
    def test_click_reveal_fails_video(self):
        deck = plain_deck()
        for index in (1, 2):
            deck[index] = reveal_slide()
        code, data = self.check(deck, "--muc", "vua", "--video")
        self.assertEqual(code, 1)
        self.assertEqual(data["error"]["step"], "video")
        self.assertIn("trang 2, 3", data["error"]["message"])

    def test_flip_card_fails_video(self):
        deck = plain_deck()
        deck[1] = reveal_slide(2, "afterEffect")
        deck[2] = slide(timing_xml=timing(interactive_seq()))
        code, data = self.check(deck, "--muc", "vua", "--video")
        self.assertEqual(code, 1)
        self.assertIn("ô bấm", data["error"]["message"])

    def test_auto_running_effects_and_morph_pass_video(self):
        deck = plain_deck()
        deck[1] = reveal_slide(3, "afterEffect")
        deck[2] = reveal_slide(2, "withEffect")
        deck[3] = slide(transition=MORPH)
        code, data = self.check(deck, "--muc", "nhieu", "--video")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["max_click_steps"], 0)

    def test_video_without_level_checks_only_clicks(self):
        code, data = self.check(plain_deck(), "--video")
        self.assertEqual(code, 0, data)
        deck = plain_deck()
        deck[1] = reveal_slide(1)
        code, data = self.check(deck, "--video")
        self.assertEqual((code, data["error"]["step"]), (1, "video"))

    def test_level_error_is_reported_before_video_error(self):
        deck = plain_deck()
        deck[1] = reveal_slide(1)
        code, data = self.check(deck, "--muc", "nhieu", "--video")
        self.assertEqual(data["error"]["step"], "muc")


class CommonWarningTest(CheckerCase):
    QUIZ = ("Câu 1. SO2 là oxide gì?", "A. base", "B. acid", "C. lưỡng tính", "D. trung tính")

    def test_quiz_without_reveal_or_answer_slide_warns(self):
        deck = plain_deck()
        deck[2] = slide(lines=self.QUIZ)
        _, data = self.check(deck, "--muc", "khong")
        self.assertIn("Trang 3 có câu trắc nghiệm nhưng chưa có cách hiện đáp án.", data["warnings"])

    def test_quiz_followed_by_answer_slide_does_not_warn(self):
        deck = plain_deck()
        deck[2] = slide(lines=self.QUIZ)
        deck[3] = slide(lines=("ĐÁP ÁN: B",))
        _, data = self.check(deck, "--muc", "khong")
        self.assertEqual(data["warnings"], [])

    def test_outline_letters_are_not_a_quiz(self):
        deck = plain_deck()
        deck[1] = slide(lines=("Nội dung báo cáo", "A. Đặt vấn đề", "B. Kết quả", "C. Phương hướng"))
        _, data = self.check(deck, "--muc", "khong")
        self.assertEqual(data["warnings"], [])

    def test_decomposed_vietnamese_answer_heading_is_recognised(self):
        import unicodedata

        deck = plain_deck()
        deck[2] = slide(lines=self.QUIZ)
        deck[3] = slide(lines=(unicodedata.normalize("NFD", "Đáp án: B"),))
        _, data = self.check(deck, "--muc", "khong")
        self.assertEqual(data["warnings"], [])

    def test_title_fade_alone_does_not_count_as_answer_reveal(self):
        deck = plain_deck()
        deck[2] = slide(timing_xml=timing(main_seq(effect("afterEffect"))), lines=self.QUIZ)
        _, data = self.check(deck, "--muc", "vua")
        self.assertIn("Trang 3 có câu trắc nghiệm nhưng chưa có cách hiện đáp án.", data["warnings"])

    def test_quiz_with_reveal_does_not_warn(self):
        deck = plain_deck()
        deck[2] = slide(timing_xml=timing(interactive_seq()), lines=self.QUIZ)
        _, data = self.check(deck, "--muc", "vua")
        self.assertFalse(any("trắc nghiệm" in w for w in data["warnings"]))

    def test_long_effect_warns(self):
        deck = plain_deck()
        deck[1] = slide(timing_xml=timing(main_seq(effect("afterEffect", dur=2500))))
        _, data = self.check(deck, "--muc", "vua")
        self.assertIn("Trang 2 có hiệu ứng dài quá 2 giây.", data["warnings"])


class InputTest(CheckerCase):
    def test_missing_level_is_input_error(self):
        code, data = self.check(plain_deck())
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_unknown_level_is_input_error(self):
        code, data = self.check(plain_deck(), "--muc", "vừa")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_missing_file_is_input_error(self):
        code, data = self.run_tool(self.folder / "khong-co.pptx", "--muc", "vua")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_not_a_zip_is_parse_error(self):
        path = self.folder / "hong.pptx"
        path.write_text("không phải pptx", encoding="utf-8")
        code, data = self.run_tool(path, "--muc", "vua")
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))

    def test_zip_without_presentation_is_parse_error(self):
        path = self.folder / "rong.pptx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", "<w/>")
        code, data = self.run_tool(path, "--muc", "vua")
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))
        self.assertIn("presentation.xml", data["error"]["message"])

    def test_broken_slide_xml_is_parse_error(self):
        deck = plain_deck()
        deck[1] = "<p:sld"
        code, data = self.check(deck, "--muc", "vua")
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))

    def test_unexpected_exception_still_prints_one_json_line(self):
        with mock.patch.object(kiem_hieu_ung, "read_deck", side_effect=RuntimeError("boom")):
            code, data = self.check(plain_deck(), "--muc", "vua")
        self.assertEqual((code, data["error"]["step"]), (1, "internal"))
        self.assertIn("boom", data["error"]["message"])


if __name__ == "__main__":
    unittest.main()
