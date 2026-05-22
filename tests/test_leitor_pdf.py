import unittest

from leitor_pdf import (
    COLUMN_RANGES,
    OUTPUT_COLUMNS,
    ExtractedLine,
    PaymentComponent,
    align_payment_sequences,
    build_output_row,
    build_precatorio_number,
    count_auto_adjustment_rows,
    count_manual_review_rows,
    extract_payment_types,
    extract_process_id,
    is_structural_line,
    join_name_fragments,
    normalize_ascii,
    reduce_payment_components,
)


class LeitorPdfRulesTest(unittest.TestCase):
    def test_extract_process_id_keeps_name_overflow(self) -> None:
        process_id, remainder = extract_process_id("202300162048 LEA")
        self.assertEqual(process_id, "202300162048")
        self.assertEqual(remainder, "LEA")

    def test_extract_payment_types_supports_partial_and_cession(self) -> None:
        types = extract_payment_types(
            [
                "Parcial Antecipação",
                "Pagamento Integral",
                "Cessão - Acordo Direto",
            ]
        )
        self.assertEqual(
            types,
            [
                "Parcial Antecipação",
                "Pagamento Integral",
                "Cessão - Acordo Direto",
            ],
        )

    def test_build_precatorio_number_applies_business_rule(self) -> None:
        self.assertEqual(build_precatorio_number("201900120859"), "12085919")
        self.assertEqual(build_precatorio_number("202000131717"), "13171720")
        self.assertEqual(build_precatorio_number("201700115823"), "11582317")

    def test_build_precatorio_number_returns_empty_for_invalid_process(self) -> None:
        self.assertEqual(build_precatorio_number(""), "")
        self.assertEqual(build_precatorio_number("20190012085"), "")
        self.assertEqual(build_precatorio_number("2019A0120859"), "")

    def test_align_payment_sequences_discards_duplicate_tail(self) -> None:
        payment_types = [
            "Pagamento Antecipação",
            "Pagamento Integral",
            "Cessão - Acordo",
        ]
        dates = [
            "03/05/2024",
            "10/05/2024",
            "10/05/2024",
            "10/05/2024",
        ]
        values = [
            "38.930,10",
            "37.410,60",
            "24.940,40",
            "37.410,60",
        ]

        aligned_types, aligned_dates, aligned_values = align_payment_sequences(
            payment_types,
            dates,
            values,
        )

        self.assertEqual(aligned_types, payment_types)
        self.assertEqual(aligned_dates, ["03/05/2024", "10/05/2024", "10/05/2024"])
        self.assertEqual(aligned_values, ["38.930,10", "37.410,60", "24.940,40"])

    def test_reduce_payment_components_keeps_integral_without_cession(self) -> None:
        payment_type, payment_date, gross_total, retained_count, had_cession = reduce_payment_components(
            [
                PaymentComponent("Pagamento Integral", "18/09/2024", 421051.33),
                PaymentComponent("Cessão - Acordo Direto", "18/09/2024", 280700.88),
            ]
        )

        self.assertEqual(payment_type, "Pagamento Integral")
        self.assertEqual(payment_date, "18/09/2024")
        self.assertEqual(gross_total, 421051.33)
        self.assertEqual(retained_count, 1)
        self.assertTrue(had_cession)

    def test_reduce_payment_components_sums_anticipation_and_integral(self) -> None:
        payment_type, payment_date, gross_total, retained_count, had_cession = reduce_payment_components(
            [
                PaymentComponent("Pagamento Antecipação", "02/04/2024", 38930.10),
                PaymentComponent("Pagamento Integral", "16/10/2024", 541197.27),
                PaymentComponent("Cessão - Acordo", "16/10/2024", 360798.18),
            ]
        )

        self.assertEqual(payment_type, "Pagamento Antecipação + Pagamento Integral")
        self.assertEqual(payment_date, "02/04/2024 | 16/10/2024")
        self.assertEqual(gross_total, 580127.37)
        self.assertEqual(retained_count, 2)
        self.assertTrue(had_cession)

    def test_join_name_fragments_preserves_line_order_with_process_overflow(self) -> None:
        record = {
            "processo": ["200800102636", "LIMA"],
            "credor": [
                "JULIO CESAR DOS SANTOS AMOROSO DE",
                "HERDEIRO DE GERMINIO AMOROSO DE",
                "LIMA",
            ],
            "__name_segments__": [
                "JULIO CESAR DOS SANTOS AMOROSO DE",
                "LIMA HERDEIRO DE GERMINIO AMOROSO DE",
                "LIMA",
            ],
        }

        self.assertEqual(
            join_name_fragments(record, "LIMA"),
            "JULIO CESAR DOS SANTOS AMOROSO DE LIMA HERDEIRO DE GERMINIO AMOROSO DE LIMA",
        )

    def test_join_name_fragments_merges_process_overflow_between_name_lines(self) -> None:
        record = {
            "processo": ["200800102636", "DOS"],
            "credor": [
                "ANNE GRAZIELLY DOS SANTOS MELO REP",
                "POR SUA GENITORA MARIA ROSANGELA",
                "SANTOS- HERDEIRA DE JOSE DELFINO",
                "DE MELO",
            ],
            "__name_segments__": [
                "ANNE GRAZIELLY DOS SANTOS MELO REP",
                "POR SUA GENITORA MARIA ROSANGELA",
                "DOS SANTOS- HERDEIRA DE JOSE DELFINO",
                "DE MELO",
            ],
        }

        self.assertEqual(
            join_name_fragments(record, "DOS"),
            "ANNE GRAZIELLY DOS SANTOS MELO REP POR SUA GENITORA MARIA ROSANGELA DOS SANTOS- HERDEIRA DE JOSE DELFINO DE MELO",
        )

    def test_build_output_row_fills_precatorio_from_process(self) -> None:
        record = {
            "processo": ["201900120859"],
            "credor": ["CREDOR EXEMPLO"],
            "entidade": ["ENTE EXEMPLO"],
            "data_pagamento": ["15/04/2024"],
            "tipo_pagamento": ["Pagamento Integral"],
            "tipo_antecipacao": [],
            "valor_bruto": ["1.000,00"],
            "previdencia": [],
            "imposto_renda": [],
            "bloqueio": [],
            "valor_liquido": ["900,00"],
            "__name_segments__": ["CREDOR EXEMPLO"],
        }

        built_row = build_output_row(record)

        self.assertEqual(built_row.row["Nº do processo"], "201900120859")
        self.assertEqual(built_row.row["Nº precatorio"], "12085919")


    def test_build_output_row_infers_anticipation_from_orphan_payment_token(self) -> None:
        record = {
            "processo": ["201300105009"],
            "credor": [
                "MARIA APARECIDA DOS SANTOS - MEEIRA",
                "DE JOSE RINALDO BEZERRA ARAUJO -",
                "FALECIDO) E HERDEIRA DE JOSEFA",
                "BEZERRA ARAUJO",
            ],
            "entidade": ["ESTADO DE SERGIPE"],
            "data_pagamento": ["24/11/2025", "24/11/2025", "27/11/2025"],
            "tipo_pagamento": ["Pagamento Integral", "CessÃ£o - Acordo Direto", "Pagamento"],
            "tipo_antecipacao": ["-"],
            "valor_bruto": ["3.333,36", "2.222,23", "40.787,05"],
            "previdencia": [],
            "imposto_renda": ["0,00", "0,00"],
            "bloqueio": ["666,67", "666,67", "7.831,11", "7.831,11"],
            "valor_liquido": ["25.493,37"],
            "__name_segments__": [
                "MARIA APARECIDA DOS SANTOS - MEEIRA",
                "DE JOSE RINALDO BEZERRA ARAUJO -",
                "FALECIDO) E HERDEIRA DE JOSEFA",
                "BEZERRA ARAUJO",
            ],
        }

        built_row = build_output_row(record)

        self.assertEqual(
            built_row.row["Tipo de pagamento"],
            "Pagamento Integral + Pagamento Antecipação",
        )
        self.assertEqual(built_row.row["Data pagamento"], "24/11/2025 | 27/11/2025")
        self.assertEqual(built_row.row["Valor Bruto do Pagamento"], 44120.41)

    def test_is_structural_line_flags_pdf_summary_footer(self) -> None:
        text = "Total de registros: 3715 Valor total: R$ 381.538.524,19"
        line = ExtractedLine(
            y=532.97,
            columns={key: "" for key in COLUMN_RANGES},
            text=text,
            normalized_text=normalize_ascii(text),
        )

        self.assertTrue(is_structural_line(line, top_limit=700.0))

    def test_review_counters_distinguish_manual_review_from_auto_adjustment(self) -> None:
        auto_adjustment_record = {
            "processo": ["201300105009"],
            "credor": [
                "MARIA APARECIDA DOS SANTOS - MEEIRA",
                "DE JOSE RINALDO BEZERRA ARAUJO -",
                "FALECIDO) E HERDEIRA DE JOSEFA",
                "BEZERRA ARAUJO",
            ],
            "entidade": ["ESTADO DE SERGIPE"],
            "data_pagamento": ["24/11/2025", "24/11/2025", "27/11/2025"],
            "tipo_pagamento": ["Pagamento Integral", "Cessão - Acordo Direto", "Pagamento"],
            "tipo_antecipacao": ["-"],
            "valor_bruto": ["3.333,36", "2.222,23", "40.787,05"],
            "previdencia": [],
            "imposto_renda": ["0,00", "0,00"],
            "bloqueio": ["666,67", "666,67", "7.831,11", "7.831,11"],
            "valor_liquido": ["25.493,37"],
            "__name_segments__": [
                "MARIA APARECIDA DOS SANTOS - MEEIRA",
                "DE JOSE RINALDO BEZERRA ARAUJO -",
                "FALECIDO) E HERDEIRA DE JOSEFA",
                "BEZERRA ARAUJO",
            ],
        }
        manual_review_record = {
            "processo": ["202500147025"],
            "credor": ["JAIRTON SANTOS DE ANDRADE"],
            "entidade": ["ESTADO DE SERGIPE"],
            "data_pagamento": ["23/09/2025"],
            "tipo_pagamento": ["Pagamento Antecipação"],
            "tipo_antecipacao": ["Idoso"],
            "valor_bruto": ["40.787,05"],
            "previdencia": [],
            "imposto_renda": ["0,00"],
            "bloqueio": [],
            "valor_liquido": ["40.787,05", "381.538.524,19"],
            "__name_segments__": ["JAIRTON SANTOS DE ANDRADE"],
        }

        auto_adjustment_row = build_output_row(auto_adjustment_record)
        manual_review_row = build_output_row(manual_review_record)

        self.assertEqual(auto_adjustment_row.row[OUTPUT_COLUMNS[1]], "201300105009")
        self.assertEqual(manual_review_row.row[OUTPUT_COLUMNS[1]], "202500147025")
        self.assertEqual(count_manual_review_rows([auto_adjustment_row, manual_review_row]), 1)
        self.assertEqual(count_auto_adjustment_rows([auto_adjustment_row, manual_review_row]), 1)


if __name__ == "__main__":
    unittest.main()
