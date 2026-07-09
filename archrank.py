import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="KEAM 2026 B.Arch Rank Generator", layout="wide")

st.title("KEAM 2026 B.Arch Rank List Generator")

st.write("""
Upload

1. Candidate Master Excel
2. NATA & Qualifying Examination Excel
""")

candidate_file = st.file_uploader(
    "Candidate Master",
    type=["xlsx", "xls"],
    key="candidate"
)

marks_file = st.file_uploader(
    "NATA Marks",
    type=["xlsx", "xls"],
    key="marks"
)


def find_math_mark(row):
    for i in range(1, 10):
        sub = f"SUB{i}_NAME"
        mark = f"SUB{i}_MARK"
        maxm = f"SUB{i}_MAX"

        if sub in row.index:
            if pd.notna(row[sub]):
                if str(row[sub]).strip().upper() == "MATHEMATICS":
                    return pd.Series({
                        "MATH_MARK": row[mark],
                        "MATH_MAX": row[maxm] if maxm in row.index else np.nan
                    })
    return pd.Series({"MATH_MARK": np.nan, "MATH_MAX": np.nan})


if candidate_file and marks_file:

    candidate = pd.read_excel(candidate_file)
    marks = pd.read_excel(marks_file)

    candidate.columns = candidate.columns.str.strip()
    marks.columns = marks.columns.str.strip()

    # Rename for merge
    if "ApplNo" in candidate.columns:
        candidate.rename(columns={"ApplNo": "APPLNO"}, inplace=True)

    # Required columns
    required_candidate = ["APPLNO", "Name", "DOB"]
    required_marks = [
        "APPLNO",
        "NATA_SCORE",
        "TOTALMARK",
        "TOTALMAXMARK"
    ]

    miss1 = [c for c in required_candidate if c not in candidate.columns]
    miss2 = [c for c in required_marks if c not in marks.columns]

    if miss1:
        st.error("Missing columns in Candidate File")
        st.write(miss1)
        st.stop()

    if miss2:
        st.error("Missing columns in Marks File")
        st.write(miss2)
        st.stop()

    # Merge
    df = marks.merge(
        candidate,
        on="APPLNO",
        how="left"
    )

    # Numeric conversion
    for c in [
        "NATA_SCORE",
        "TOTALMARK",
        "TOTALMAXMARK"
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Mathematics mark + its max mark, normalized to a common /100 scale
    # (subjects are not all out of the same total, e.g. 200 vs 100,
    # so raw marks cannot be compared directly for the tie-break)
    math_df = df.apply(find_math_mark, axis=1)
    df["MATH_MARK"] = pd.to_numeric(math_df["MATH_MARK"], errors="coerce").round(4)
    df["MATH_MAX"] = pd.to_numeric(math_df["MATH_MAX"], errors="coerce").round(4)
    df["MATH_PERCENT"] = (
        df["MATH_MARK"] / df["MATH_MAX"] * 100
    ).round(4)

    # DOB
    df["DOB"] = pd.to_datetime(
        df["DOB"],
        dayfirst=True,
        errors="coerce"
    )

    # Qualifying Score out of 200
    df["QUALIFY_SCORE"] = (
        df["TOTALMARK"] /
        df["TOTALMAXMARK"] *
        200
    ).round(2)

    # Final Score out of 400
    df["FINAL_SCORE"] = (
        df["NATA_SCORE"] +
        df["QUALIFY_SCORE"]
    ).round(2)

    # Sort according to prospectus
    df = df.sort_values(
        by=[
            "FINAL_SCORE",
            "NATA_SCORE",
            "MATH_PERCENT",
            "DOB",
            "APPLNO"
        ],
        ascending=[
            False,
            False,
            False,
            True,
            True
        ]
    ).reset_index(drop=True)

    df["RANK"] = range(1, len(df) + 1)

    st.success(f"Total Candidates : {len(df)}")

    display_columns = [
        "RANK",
        "APPLNO",
        "Name",
        "DOB",
        "NATA_SCORE",
        "MATH_MARK",
        "MATH_MAX",
        "MATH_PERCENT",
        "TOTALMARK",
        "TOTALMAXMARK",
        "QUALIFY_SCORE",
        "FINAL_SCORE"
    ]

    extra_cols = [
        "Category",
        "Community",
        "Gender",
        "Nativity"
    ]

    for col in extra_cols:
        if col in df.columns:
            display_columns.append(col)

    st.dataframe(
        df[display_columns],
        use_container_width=True
    )

    # Statistics
    st.subheader("Statistics")

    c1, c2, c3 = st.columns(3)

    c1.metric("Candidates", len(df))
    c2.metric("Highest Score", round(df["FINAL_SCORE"].max(), 2))
    c3.metric("Lowest Score", round(df["FINAL_SCORE"].min(), 2))

    # Excel Output
    # Create Output DataFrame
    output_df = pd.DataFrame({
        "Rank": df["RANK"],
        "ApplNo": df["APPLNO"],
        "Candidate": df["Name"],
        "NATA Score": df["NATA_SCORE"],
        "Math Score": df["MATH_MARK"],
        "Math Max": df["MATH_MAX"],
        "Math %": df["MATH_PERCENT"],
        "TOTALMARK": df["TOTALMARK"],
        "TOTALMAXMARK": df["TOTALMAXMARK"],
        "Qualifying Score (/200)": df["QUALIFY_SCORE"].round(2),
        "DOB": df["DOB"].dt.strftime("%d-%m-%Y"),
        "Final Score": df["FINAL_SCORE"]
    })

# Export to Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="BArch Rank List")

    st.dataframe(output_df, use_container_width=True)

    st.download_button(
        label="📥 Download Rank List",
        data=output.getvalue(),
        file_name="KEAM2026_BARCH_RANKLIST.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Download Rank List",
        data=output.getvalue(),
        file_name="KEAM2026_BARCH_RANKLIST.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
