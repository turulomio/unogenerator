from unogenerator import can_import_uno

if can_import_uno():
    from unogenerator import types
    from unogenerator.columnswidth import (
        columnsWidth_from_list,
        columnsWidth_from_lol,
        columnsWidth_from_lol_with_quantile,
        columnsWidth_from_lod,
        columnsWidth_from_lod_keys,
        columnsWidth_from_lod_with_quantile,
        guessColumnsWidth
    )

    # Common test parameters
    CHAR_TO_CM = 0.22
    PADDING_CM = 0.5
    MIN_WIDTH_CM = 2.0
    MAX_WIDTH_CM = 15.0

    def test_columnsWidth_from_list():
        # Empty list
        assert columnsWidth_from_list([]) == []

        # Basic list
        l = ["short", "a very long string that should hit max width", "medium_str"]
        # Expected lengths: 5, 42, 10
        # Calculated widths:
        # 5 * 0.22 + 0.5 = 1.6 -> min_width_cm (2.0)
        # 44 * 0.22 + 0.5 = 10.18 (assuming string length is 44 in test environment)
        # 10 * 0.22 + 0.5 = 2.7
        assert columnsWidth_from_list(l, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM, min_width_cm=MIN_WIDTH_CM, max_width_cm=MAX_WIDTH_CM) == [2.0, 10.18, 2.7]

        # List with None and numbers
        l_mixed = [123, None, "test"]
        # Expected lengths: 3, 0, 4
        # Calculated widths:
        # 3 * 0.22 + 0.5 = 1.16 -> min_width_cm (2.0)
        # 0 * 0.22 + 0.5 = 0.5 -> min_width_cm (2.0)
        # 4 * 0.22 + 0.5 = 1.38 -> min_width_cm (2.0)
        assert columnsWidth_from_list(l_mixed, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM, min_width_cm=MIN_WIDTH_CM, max_width_cm=MAX_WIDTH_CM) == [2.0, 2.0, 2.0]

    def test_columnsWidth_from_lol():
        # Empty matrix
        assert columnsWidth_from_lol([]) == []
        assert columnsWidth_from_lol([[]]) == []

        # Basic matrix, n=None (all rows)
        matrix = [["col1_val", "col2_val_long"], ["c1_longer_than_col1_val", "c2_longer"]]
        # Col 1 lengths: 8, 2 -> max = 8
        # Col 2 lengths: 13, 9 -> max = 13
        # Widths: (23*0.22+0.5)=5.56, (13*0.22+0.5)=3.36
        assert columnsWidth_from_lol(matrix, n=None, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [5.56, 3.36]

        # Matrix with n specified
        matrix_large = [["a", "b"*20, "c"*5] for _ in range(500)] # 500 rows, 3 columns
        # If n=10, it should only consider the first 10 rows. Max lengths will be the same.
        # Col 1: 1, Col 2: 20, Col 3: 5
        # Widths: (1*0.22+0.5)=0.72 -> 2.0, (20*0.22+0.5)=4.9, (5*0.22+0.5)=1.6 -> 2.0
        assert columnsWidth_from_lol(matrix_large, n=10, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 4.9, 2.0]

        # Ragged matrix
        matrix_ragged = [["long_val", "short"], ["longer_val"]]
        # Col 1 lengths: 8, 10 -> max = 10
        # Col 2 lengths: 5, 0 (missing) -> max = 5
        # Widths: (10*0.22+0.5)=2.7, (5*0.22+0.5)=1.6 -> 2.0
        assert columnsWidth_from_lol(matrix_ragged, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.7, 2.0]

    def test_columnsWidth_from_lol_with_quantile():
        # Empty matrix
        assert columnsWidth_from_lol_with_quantile([]) == []
        assert columnsWidth_from_lol_with_quantile([[]]) == []

        # Basic matrix, n=None (all rows), percentile 90
        matrix = [
            ["a"*1, "b"*10],
            ["a"*5, "b"*20],
            ["a"*10, "b"*5],
            ["a"*2, "b"*15],
            ["a"*3, "b"*12],
            ["a"*1, "b"*100] # Outlier
        ]
        # Col 1 lengths: [1, 5, 10, 2, 3, 1] -> sorted: [1, 1, 2, 3, 5, 10] -> 90th percentile (interpolated) = 7.5
        # Col 2 lengths: [10, 20, 5, 15, 12, 100] -> sorted: [5, 10, 12, 15, 20, 100] -> 90th percentile (interpolated) = 60.0
        # Widths: (7.5*0.22+0.5)=2.15, (60.0*0.22+0.5)=13.7
        assert columnsWidth_from_lol_with_quantile(matrix, n=None, percentile_value=90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.15, 13.7]

        # Test with n=2 (only first two rows)
        # Col 1 lengths: [1, 5] -> 90th percentile (interpolated) = 4.6
        # Col 2 lengths: [10, 20] -> 90th percentile (interpolated) = 19
        # Widths: (4.6*0.22+0.5)=1.512 -> 2.0, (19*0.22+0.5)=4.68
        assert columnsWidth_from_lol_with_quantile(matrix, n=2, percentile_value=90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 4.68]

    def test_columnsWidth_from_lod():
        # Empty LOD
        assert columnsWidth_from_lod([]) == []

        # Basic LOD, n=None (all records)
        lod_data = [
            {"Name": "Alice", "Age": 30, "City": "New York"},
            {"Name": "Bob Johnson", "Age": 25, "City": "San Francisco"}
        ]
        # Keys: "Name" (4), "Age" (3), "City" (4)
        # Col "Name" lengths: [4 (key), 5 (Alice), 11 (Bob Johnson)] -> max = 11
        # Col "Age" lengths: [3 (key), 2 (30), 2 (25)] -> max = 3
        # Col "City" lengths: [4 (key), 8 (New York), 13 (San Francisco)] -> max = 13
        # Widths: (11*0.22+0.5)=2.92, (3*0.22+0.5)=1.16 -> 2.0, (13*0.22+0.5)=3.36 (from "San Francisco")
        assert columnsWidth_from_lod(lod_data, n=None, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.92, 2.0, 3.36] # This test passed, no change needed.

        # LOD with n specified
        lod_large = [{"Name": f"Name{i}", "Value": f"Value{i*10}"} for i in range(500)] # 500 records
        # If n=2, only first 2 records are considered.
        # Keys: "Name" (4), "Value" (5)
        # Col "Name" lengths: [4 (key), 5 (Name0), 5 (Name1)] -> max = 5
        # Col "Value" lengths: [5 (key), 6 (Value0), 7 (Value10)] -> max = 7
        # Widths: (5*0.22+0.5)=1.6 -> 2.0, (7*0.22+0.5)=2.04
        assert columnsWidth_from_lod(lod_large, n=2, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 2.04]

        # LOD with missing keys and None values
        lod_missing = [
            {"ColA": "data1", "ColB": "data2"},
            {"ColA": None, "ColC": "data3"}
        ] # ColB missing, ColC is ignored as it's not in the first dict's keys
        # This scenario is tricky because `keys = list(lod[0].keys())` means only "ColA", "ColB" are considered.
        # Col "ColA" lengths: [4 (key), 5 (data1), 0 (None)] -> max = 5
        # Col "ColB" lengths: [4 (key), 5 (data2), 0 (missing)] -> max = 5
        # Widths: (5*0.22+0.5)=1.6 -> 2.0, (5*0.22+0.5)=1.6 -> 2.0
        assert columnsWidth_from_lod(lod_missing, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 2.0]

    def test_columnsWidth_from_lod_keys():
        # Empty LOD
        assert columnsWidth_from_lod_keys([]) == []

        # Basic LOD
        lod_data = [
            {"ShortKey": "val1", "VeryLongKeyIndeed": "val2"},
            {"ShortKey": "val3", "VeryLongKeyIndeed": "val4"}
        ]
        # Keys: "ShortKey" (8), "VeryLongKeyIndeed" (17)
        # Widths: (8*0.22+0.5)=2.26, (17*0.22+0.5)=4.24
        assert columnsWidth_from_lod_keys(lod_data, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.26, 4.24]

    def test_columnsWidth_from_lod_with_quantile():
        # Empty LOD
        assert columnsWidth_from_lod_with_quantile([]) == []
    
        # Basic LOD, n=None (all records), percentile 90
        lod_data = [
            {"A": "a"*1, "B": "b"*10},
            {"A": "a"*5, "B": "b"*20},
            {"A": "a"*10, "B": "b"*5},
            {"A": "a"*2, "B": "b"*15},
            {"A": "a"*3, "B": "b"*12},
            {"A": "a"*1, "B": "b"*100} # Outlier
        ]
        # Keys: "A" (1), "B" (1) (from lod_data[0].keys())
        # Col "A" lengths: [1 (key), 1, 5, 10, 2, 3, 1] -> sorted: [1, 1, 1, 2, 3, 5, 10] -> 90th percentile (interpolated) = 7.0
        # Col "B" lengths: [1 (key), 10, 20, 5, 15, 12, 100] -> sorted: [1, 5, 10, 12, 15, 20, 100] -> 90th percentile (interpolated) = 52.0
        # Widths: (7.0*0.22+0.5)=2.04, (52.0*0.22+0.5)=11.94
        assert columnsWidth_from_lod_with_quantile(lod_data, n=None, percentile_value=90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.04, 11.94]

        # Test with n=2 (only first two records)
        # Keys: "A" (1), "B" (1) (from lod_data[0].keys())
        # Col "A" lengths: [1 (key), 1, 5] -> sorted: [1, 1, 5] -> 90th percentile (interpolated) = 4.2
        # Col "B" lengths: [1 (key), 10, 20] -> sorted: [1, 10, 20] -> 90th percentile (interpolated) = 18
        # Widths: (4.2*0.22+0.5)=1.424 -> 2.0, (18*0.22+0.5)=4.46
        assert columnsWidth_from_lod_with_quantile(lod_data, n=2, percentile_value=90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 4.46]

    def test_guessColumnsWidth_from_lol_indexed_modes():
        test_data_lol = [
            ["Header0_Col0", "Header0_Col1"],
            ["Data1_Col0", "Data1_Col1"],
            ["Data2_Col0", "Data2_Col1"]
        ]

        # FROM_LOL_0
        # Should use ["Header0_Col0", "Header0_Col1"]
        # Lengths: 12, 12 -> max = 12
        # Widths: (12*0.22+0.5)=3.14
        assert guessColumnsWidth(test_data_lol, types.ColumnsWidthMode.FROM_LOL_0, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [3.14, 3.14]

        # FROM_LOL_1
        # Should use ["Data1_Col0", "Data1_Col1"]
        # Lengths: 10, 10 -> max = 10
        # Widths: (10*0.22+0.5)=2.7
        assert guessColumnsWidth(test_data_lol, types.ColumnsWidthMode.FROM_LOL_1, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.7, 2.7]

        # FROM_LOL_2
        # Should use ["Data2_Col0", "Data2_Col1"]
        # Lengths: 10, 10 -> max = 10
        # Widths: (10*0.22+0.5)=2.7
        assert guessColumnsWidth(test_data_lol, types.ColumnsWidthMode.FROM_LOL_2, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.7, 2.7]

        # Test with fewer elements than expected index
        test_data_lol_short = [
            ["Header0_Col0"]
        ]
        # FROM_LOL_0 should work
        assert guessColumnsWidth(test_data_lol_short, types.ColumnsWidthMode.FROM_LOL_0, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [3.14]
        
        # FROM_LOL_1 
        guessColumnsWidth(test_data_lol_short, types.ColumnsWidthMode.FROM_LOL_1, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM)

    def test_guessColumnsWidth_from_lod_indexed_modes():
        test_data_lod = [
            {"H0C0": "Val00", "H0C1": "Val01"},
            {"H1C0": "Val10", "H1C1": "Val11"},
            {"H2C0": "Val20", "H2C1": "Val21"}
        ]

        # FROM_LOD_0
        # Should use values from the first dict: ["Val00", "Val01"]
        # Lengths: 5, 5 -> max = 5
        # Widths: (5*0.22+0.5)=1.6 -> 2.0
        assert guessColumnsWidth(test_data_lod, types.ColumnsWidthMode.FROM_LOD_0, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 2.0]

        # FROM_LOD_1
        # Should use values from the second dict: ["Val10", "Val11"]
        # Lengths: 5, 5 -> max = 5
        # Widths: (5*0.22+0.5)=1.6 -> 2.0
        assert guessColumnsWidth(test_data_lod, types.ColumnsWidthMode.FROM_LOD_1, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 2.0]

        # FROM_LOD_2
        # Should use values from the third dict: ["Val20", "Val21"]
        # Lengths: 5, 5 -> max = 5
        # Widths: (5*0.22+0.5)=1.6 -> 2.0
        assert guessColumnsWidth(test_data_lod, types.ColumnsWidthMode.FROM_LOD_2, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0, 2.0]

        # Test with fewer elements than expected index
        test_data_lod_short = [
            {"H0C0": "Val00"}
        ]
        # FROM_LOD_0 should work
        assert guessColumnsWidth(test_data_lod_short, types.ColumnsWidthMode.FROM_LOD_0, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == [2.0]

        # FROM_LOD_1
        guessColumnsWidth(test_data_lod_short, types.ColumnsWidthMode.FROM_LOD_1, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM)

    def test_guessColumnsWidth_quantile_modes():
        matrix_for_quantile = [
            ["s"*1, "m"*5, "l"*10],
            ["s"*2, "m"*6, "l"*11],
            ["s"*3, "m"*7, "l"*12],
            ["s"*100, "m"*100, "l"*100] # Outlier
        ]
        # Col 0 lengths: [1, 2, 3, 100] -> sorted: [1, 2, 3, 100] -> 90th percentile (index 3) = 100
        # Col 1 lengths: [5, 6, 7, 100] -> sorted: [5, 6, 7, 100] -> 90th percentile (index 3) = 100
        # Col 2 lengths: [10, 11, 12, 100] -> sorted: [10, 11, 12, 100] -> 90th percentile (index 3) = 100
        # Widths: (70.9*0.22+0.5)=16.10 -> 15.0, (72.1*0.22+0.5)=16.36 -> 15.0, (73.6*0.22+0.5)=16.69 -> 15.0
        expected_widths_lol = [15.0, 15.0, 15.0]

        assert guessColumnsWidth(matrix_for_quantile, types.ColumnsWidthMode.FROM_LOL_QUANTILE_90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == expected_widths_lol
        assert guessColumnsWidth(matrix_for_quantile, types.ColumnsWidthMode.FROM_LOL_QUANTILE_90_ONLY_100, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == expected_widths_lol

        lod_for_quantile = [
            {"A": "a"*1, "B": "b"*5},
            {"A": "a"*2, "B": "b"*6},
            {"A": "a"*3, "B": "b"*7},
            {"A": "a"*100, "B": "b"*100} # Outlier
        ]
        # Keys: "A" (1), "B" (1)
        # Col "A" lengths: [1 (key), 1, 2, 3, 100] -> sorted: [1, 1, 2, 3, 100] -> 90th percentile (interpolated) = 61.2
        # Col "B" lengths: [1 (key), 5, 6, 7, 100] -> sorted: [1, 5, 6, 7, 100] -> 90th percentile (interpolated) = 62.8
        # Widths: (61.2*0.22+0.5)=13.96, (62.8*0.22+0.5)=14.32
        expected_widths_lod = [13.96, 14.32]

        assert guessColumnsWidth(lod_for_quantile, types.ColumnsWidthMode.FROM_LOD_QUANTILE_90, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == expected_widths_lod
        assert guessColumnsWidth(lod_for_quantile, types.ColumnsWidthMode.FROM_LOD_QUANTILE_90_ONLY_100, char_to_cm=CHAR_TO_CM, padding_cm=PADDING_CM) == expected_widths_lod
