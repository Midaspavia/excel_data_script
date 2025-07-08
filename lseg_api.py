import refinitiv.data as rd

def fetch_lseg_data(ric: str, fields: list) -> dict:
    try:
        rd.open_session()

        response = rd.Content.FundamentalAndReference.Definition(
            universe=[ric],
            fields=fields
        ).get_data()

        if response and not response.is_success:
            print(f"Fehler bei {ric}: {response.message}")
            return {}

        df = response.data.df
        if df.empty:
            return {}

        return {field: df[field].iloc[0] for field in fields if field in df.columns}
    except Exception as e:
        print(f"Fehler bei RIC {ric}: {e}")
        return {}