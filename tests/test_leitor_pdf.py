import unittest

from leitor_pdf import (
    PaymentComponent,
    align_payment_sequences,
    build_output_row,
    build_precatorio_number,
    extract_payment_types,
    extract_process_id,
    join_name_fragments,
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


if __name__ == "__main__":
    unittest.main()
