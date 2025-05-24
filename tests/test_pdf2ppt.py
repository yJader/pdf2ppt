from pathlib import Path
import sys

# 添加项目根目录到 Python 路径
sys.path.append(str(Path(__file__).resolve().parents[1]))


from pdf2ppt import extract_pdf_comments_with_pages, convert_pdf_to_ppt_with_comments


def test_extract_pdf_comments_with_pages():
    comments = extract_pdf_comments_with_pages(
        Path(__file__).parent / "assets" / "test.pdf"
    )
    assert comments is not None
    print(comments)
    assert len(comments) > 0
    assert comments[1][0] == "这是测试的备注"


def test_convert_pdf_to_ppt_with_comments():
    comments = extract_pdf_comments_with_pages(
        Path(__file__).parent / "assets" / "test.pdf"
    )

    convert_pdf_to_ppt_with_comments(
        Path(__file__).parent / "assets" / "test.pdf",
        Path(__file__).parent / "assets" / "test.pptx",
        comments,
    )
    assert (Path(__file__).parent / "assets" / "test.pptx").exists()
