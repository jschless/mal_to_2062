import pandas as pd
from pypdftk import fill_form, concat
import os


def read_mal(path):
    df = pd.read_excel(path, dtype=str)
    df["name_and_rank"] = df["RANK"] + " " + df["NAME"]
    df["M4"] = df["BS# "] + " / " + df["M4 CARBINE"]
    df["PVS-14"] = (
        df["AD #"]
        + " / "
        + df["NVG"].apply(lambda x: "" if pd.isna(x) else str(int(x)))
    )

    def get_last_name(name):
        if not isinstance(name, str):
            return ""
        if "," in name:
            return name.split(",")[0]
        else:
            return name

    df["last_name"] = df.NAME.apply(get_last_name)
    df["name_and_rank"] = df.apply(
        lambda x: x["last_name"] if pd.isna(x["name_and_rank"]) else x["name_and_rank"],
        axis=1,
    )

    return df


def get_params(example_data):
    # define fields of dataframe I care about
    FIELDS = [
        "M4",
        "PVS-14",
        "M17",
        "M249 SAW",
        "OPTIC",
        "PEQ-15",
        "CREW PEQ-15",
        "M240",
        "M145 SCOPE",
        "AN/PAS-13 V2",
        "M320 GL",
        "M26 SHOTGUN",
        "LRF",
        "THERMAL",
        "BINOS",
        "50 CAL",
        "AN/PAS-13 V3",
        "MK 19",
        "CLU",
        "SURE FIRE",
        "Barrel Gloves",
        "Spare Barrel",
        "Barrel Handle",
        "50 BFA",
        "MG BFA",
        "PAS Mount",
        "Barrel Bag",
        "Rhino Mount",
        "J Hook",
        "Sling",
        "M4 BFA",
    ]

    example_data = {k: v for k, v in example_data.items() if not pd.isna(v)}
    # add additional equipment depending on what people are assigned
    if "50 CAL" in example_data:
        example_data["Barrel Gloves"] = ""
        example_data["Spare Barrel"] = ""
        example_data["Barrel Handle"] = ""
        example_data["50 BFA"] = ""
        example_data["PAS Mount"] = ""
    elif "M249 SAW" in example_data or "M240" in example_data:
        example_data["Spare Barrel"] = ""
        example_data["Barrel Bag"] = ""
        example_data["MG BFA"] = ""
        example_data["Sling"] = ""
    else:
        example_data["M4 BFA"] = ""
        example_data["Sling"] = ""

    example_data["Rhino Mount"] = ""
    example_data["J Hook"] = ""

    # define the order I want things to display
    ordering = {v: i for i, v in enumerate(FIELDS)}

    # I only want to fill the intersection of FIELDS and what exists in example_data
    rel_fields = set(FIELDS).intersection(example_data.keys())

    ITEM_KEY = "ITEM DESCRIPTION bRow"
    params = {
        ITEM_KEY + str(i + 1): f + f"  ({example_data[f]})"
        for i, f in enumerate(sorted(list(rel_fields), key=lambda x: ordering[x]))
    }
    params["TO"] = example_data["name_and_rank"]

    return params


def main():
    directory_path = os.getcwd()
    df = read_mal(os.path.join(directory_path, "62nd Master MAL.xlsx"))

    records = [
        r
        for r in df.to_dict("records")
        if not pd.isna(r["NAME"])
        and r["NAME"] != "DEADLINE"
        and r["NAME"] != "UNASSIGNED"
    ]
    platoons = set([r["Platoon"] for r in records])

    base_dir = os.path.join(directory_path, "2062s")
    pdf_path = os.path.join(directory_path, "2062_forms_redone.pdf")

    # creating directories
    if not os.path.exists(base_dir):
        os.mkdir(base_dir)

    for p in platoons:
        if not os.path.exists(os.path.join(base_dir, p)):
            os.mkdir(os.path.join(base_dir, p))

    # filling pdfs
    for r in records:
        try:
            out_file = os.path.join(
                base_dir, r["Platoon"], f"{r['last_name']}_2062.pdf"
            )
            params = get_params(r)
            fill_form(pdf_path, params, out_file, flatten=False)
        except Exception as e:
            print("issue with ", r, e)

    # combining pdfs
    for p in platoons:
        temp_path = os.path.join(base_dir, p)
        concat(
            files=[os.path.join(temp_path, f) for f in os.listdir(temp_path)],
            out_file=os.path.join(base_dir, f"combined_2062s_{p}.pdf"),
        )


if __name__ == "__main__":
    main()
