using ExcelDna.Integration;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms; // <- aggiungi questa using in cima al file
using System.Xml.Linq;

public class PsicroAddIn
{
    //const string DLL_PATH = @"C:\Users\Utente\source\repos\psicro\x64\Debug\psicro.dll";
    const string DLL_PATH = "psicro.dll";
    #region 1. DICHIARAZIONI DLL C (Import)


    // Carichiamo la funzione per la quota dalla tua DLL C
    [DllImport("psicro.dll", CallingConvention = CallingConvention.StdCall)]public static extern double Excel_set_quota(double altitude);
    //[DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_get_patm_at_altitude(double altitude);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_Psat(double t);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_TPsat(double p_kpa);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_xsat_t(double t);

    // T - Temperatura
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_ur_x(double ur, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_ur_h(double ur, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_ur_vau(double ur, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_ur_tbu(double ur, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_ur_tr(double ur, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_x_h(double x, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_x_vau(double x, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_x_tbu(double x, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_x_tr(double x, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_h_vau(double h, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_h_tbu(double h, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_h_tr(double h, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_vau_tbu(double vau, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_vau_tr(double vau, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_t_tbu_tr(double tbu, double tr);

    // UR - Umidità Relativa
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_t_x(double t, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_t_h(double t, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_t_vau(double t, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_t_tbu(double t, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_t_tr(double t, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_x_h(double x, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_x_vau(double x, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_x_tbu(double x, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_x_tr(double x, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_h_vau(double h, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_h_tbu(double h, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_h_tr(double h, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_vau_tbu(double vau, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_vau_tr(double vau, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_ur_tbu_tr(double tbu, double tr);

    // X - Titolo (Umidità specifica)
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_t_ur(double t, double ur);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_t_h(double t, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_t_vau(double t, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_t_tbu(double t, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_t_tr(double t, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_ur_h(double ur, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_ur_vau(double ur, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_ur_tbu(double ur, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_ur_tr(double ur, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_h_vau(double h, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_h_tbu(double h, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_h_tr(double h, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_vau_tbu(double vau, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_vau_tr(double vau, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_x_tbu_tr(double tbu, double tr);

    // H - Entalpia
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_t_ur(double t, double ur);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_t_x(double t, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_t_vau(double t, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_t_tbu(double t, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_t_tr(double t, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_ur_x(double ur, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_ur_vau(double ur, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_ur_tbu(double ur, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_ur_tr(double ur, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_x_vau(double x, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_x_tbu(double x, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_x_tr(double x, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_vau_tbu(double vau, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_vau_tr(double vau, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_h_tbu_tr(double tbu, double tr);

    // VAU - Volume specifico
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_t_ur(double t, double ur);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_t_x(double t, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_t_h(double t, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_t_tbu(double t, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_t_tr(double t, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_ur_x(double ur, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_ur_h(double ur, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_ur_tbu(double ur, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_ur_tr(double ur, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_x_h(double x, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_x_tbu(double x, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_x_tr(double x, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_h_tbu(double h, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_h_tr(double h, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_vau_tbu_tr(double tbu, double tr);

    // TBU - Bulbo umido
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_t_ur(double t, double ur);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_t_x(double t, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_t_h(double t, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_t_vau(double t, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_t_tr(double t, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_ur_x(double ur, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_ur_h(double ur, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_ur_vau(double ur, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_ur_tr(double ur, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_x_h(double x, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_x_vau(double x, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_x_tr(double x, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_h_vau(double h, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_h_tr(double h, double tr);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tbu_vau_tr(double vau, double tr);

    // TR - Punto di rugiada
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_t_ur(double t, double ur);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_t_x(double t, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_t_h(double t, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_t_vau(double t, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_t_tbu(double t, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_ur_x(double ur, double x);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_ur_h(double ur, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_ur_vau(double ur, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_ur_tbu(double ur, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_x_h(double x, double h);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_x_vau(double x, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_x_tbu(double x, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_h_vau(double h, double vau);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_h_tbu(double h, double tbu);
    [DllImport(DLL_PATH, CallingConvention = CallingConvention.StdCall)] public static extern double Excel_tr_vau_tbu(double vau, double tbu);

    #endregion

    #region 2. FUNZIONI EXCEL (Esposizione)

    private const string CAT = "Psicrometria";




    // --- COSTANTI E MAPPATURA ---
    // 0: t | 1: ur | 2: x | 3: h | 4: vau | 5: tbu | 6: tr
    //private const string CAT = "Psicrometria";

    [ExcelFunction(Name = "PSICRO.PROP", Description = "Funzione Universale: Target, Prop1, Prop2, Val1, Val2", Category = CAT)]
    public static object PsicroProp(string target, string p1, string p2, double val1, double val2)
    {
        try
        {
            // 1. Identificazione Indici
            int idT = GetPropIndex(target);
            int id1 = GetPropIndex(p1);
            int id2 = GetPropIndex(p2);

            if (idT == -1 || id1 == -1 || id2 == -1)
                return "#NOME_PROP_ERRATO# (Usa: t, ur, x, h, vau, tbu, tr)";

            if (id1 == id2)
                return "#ERRORE: Proprietà di input identiche#";

            // 2. Ordinamento per PairID (i1 sempre < i2) e swap valori
            int i1, i2;
            double v1, v2;
            if (id1 < id2)
            {
                i1 = id1; i2 = id2; v1 = val1; v2 = val2;
            }
            else
            {
                i1 = id2; i2 = id1; v1 = val2; v2 = val1;
            }

            int pairID = (i1 * 10) + i2;
            double ris = 0;

            // 3. Logica di Calcolo basata sul Target
            switch (idT)
            {
                case 0: // --- TARGET t [0] ---
                    if (pairID == 12) ris = Excel_t_ur_x(v1, v2);
                    else if (pairID == 13) ris = Excel_t_ur_h(v1, v2);
                    else if (pairID == 14) ris = Excel_t_ur_vau(v1, v2);
                    else if (pairID == 15) ris = Excel_t_ur_tbu(v1, v2);
                    else if (pairID == 16) ris = Excel_t_ur_tr(v1, v2);
                    else if (pairID == 23) ris = Excel_t_x_h(v1, v2);
                    else if (pairID == 24) ris = Excel_t_x_vau(v1, v2);
                    else if (pairID == 25) ris = Excel_t_x_tbu(v1, v2);
                    else if (pairID == 26) ris = Excel_t_x_tr(v1, v2);
                    else if (pairID == 34) ris = Excel_t_h_vau(v1, v2);
                    else if (pairID == 35) ris = Excel_t_h_tbu(v1, v2);
                    else if (pairID == 36) ris = Excel_t_h_tr(v1, v2);
                    else if (pairID == 45) ris = Excel_t_vau_tbu(v1, v2);
                    else if (pairID == 46) ris = Excel_t_vau_tr(v1, v2);
                    else if (pairID == 56) ris = Excel_t_tbu_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 1: // --- TARGET ur [1] ---
                    if (pairID == 02) ris = Excel_ur_t_x(v1, v2);
                    else if (pairID == 03) ris = Excel_ur_t_h(v1, v2);
                    else if (pairID == 04) ris = Excel_ur_t_vau(v1, v2);
                    else if (pairID == 05) ris = Excel_ur_t_tbu(v1, v2);
                    else if (pairID == 06) ris = Excel_ur_t_tr(v1, v2);
                    else if (pairID == 23) ris = Excel_ur_x_h(v1, v2);
                    else if (pairID == 24) ris = Excel_ur_x_vau(v1, v2);
                    else if (pairID == 25) ris = Excel_ur_x_tbu(v1, v2);
                    else if (pairID == 26) ris = Excel_ur_x_tr(v1, v2);
                    else if (pairID == 34) ris = Excel_ur_h_vau(v1, v2);
                    else if (pairID == 35) ris = Excel_ur_h_tbu(v1, v2);
                    else if (pairID == 36) ris = Excel_ur_h_tr(v1, v2);
                    else if (pairID == 45) ris = Excel_ur_vau_tbu(v1, v2);
                    else if (pairID == 46) ris = Excel_ur_vau_tr(v1, v2);
                    else if (pairID == 56) ris = Excel_ur_tbu_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 2: // --- TARGET x [2] ---
                    if (pairID == 01) ris = Excel_x_t_ur(v1, v2);
                    else if (pairID == 03) ris = Excel_x_t_h(v1, v2);
                    else if (pairID == 04) ris = Excel_x_t_vau(v1, v2);
                    else if (pairID == 05) ris = Excel_x_t_tbu(v1, v2);
                    else if (pairID == 06) ris = Excel_x_t_tr(v1, v2);
                    else if (pairID == 13) ris = Excel_x_ur_h(v1, v2);
                    else if (pairID == 14) ris = Excel_x_ur_vau(v1, v2);
                    else if (pairID == 15) ris = Excel_x_ur_tbu(v1, v2);
                    else if (pairID == 16) ris = Excel_x_ur_tr(v1, v2);
                    else if (pairID == 34) ris = Excel_x_h_vau(v1, v2);
                    else if (pairID == 35) ris = Excel_x_h_tbu(v1, v2);
                    else if (pairID == 36) ris = Excel_x_h_tr(v1, v2);
                    else if (pairID == 45) ris = Excel_x_vau_tbu(v1, v2);
                    else if (pairID == 46) ris = Excel_x_vau_tr(v1, v2);
                    else if (pairID == 56) ris = Excel_x_tbu_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 3: // --- TARGET h [3] ---
                    if (pairID == 01) ris = Excel_h_t_ur(v1, v2);
                    else if (pairID == 02) ris = Excel_h_t_x(v1, v2);
                    else if (pairID == 04) ris = Excel_h_t_vau(v1, v2);
                    else if (pairID == 05) ris = Excel_h_t_tbu(v1, v2);
                    else if (pairID == 06) ris = Excel_h_t_tr(v1, v2);
                    else if (pairID == 12) ris = Excel_h_ur_x(v1, v2);
                    else if (pairID == 14) ris = Excel_h_ur_vau(v1, v2);
                    else if (pairID == 15) ris = Excel_h_ur_tbu(v1, v2);
                    else if (pairID == 16) ris = Excel_h_ur_tr(v1, v2);
                    else if (pairID == 24) ris = Excel_h_x_vau(v1, v2);
                    else if (pairID == 25) ris = Excel_h_x_tbu(v1, v2);
                    else if (pairID == 26) ris = Excel_h_x_tr(v1, v2);
                    else if (pairID == 45) ris = Excel_h_vau_tbu(v1, v2);
                    else if (pairID == 46) ris = Excel_h_vau_tr(v1, v2);
                    else if (pairID == 56) ris = Excel_h_tbu_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 4: // --- TARGET vau [4] ---
                    if (pairID == 01) ris = Excel_vau_t_ur(v1, v2);
                    else if (pairID == 02) ris = Excel_vau_t_x(v1, v2);
                    else if (pairID == 03) ris = Excel_vau_t_h(v1, v2);
                    else if (pairID == 05) ris = Excel_vau_t_tbu(v1, v2);
                    else if (pairID == 06) ris = Excel_vau_t_tr(v1, v2);
                    else if (pairID == 12) ris = Excel_vau_ur_x(v1, v2);
                    else if (pairID == 13) ris = Excel_vau_ur_h(v1, v2);
                    else if (pairID == 15) ris = Excel_vau_ur_tbu(v1, v2);
                    else if (pairID == 16) ris = Excel_vau_ur_tr(v1, v2);
                    else if (pairID == 23) ris = Excel_vau_x_h(v1, v2);
                    else if (pairID == 25) ris = Excel_vau_x_tbu(v1, v2);
                    else if (pairID == 26) ris = Excel_vau_x_tr(v1, v2);
                    else if (pairID == 35) ris = Excel_vau_h_tbu(v1, v2);
                    else if (pairID == 36) ris = Excel_vau_h_tr(v1, v2);
                    else if (pairID == 56) ris = Excel_vau_tbu_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 5: // --- TARGET tbu [5] ---
                    if (pairID == 01) ris = Excel_tbu_t_ur(v1, v2);
                    else if (pairID == 02) ris = Excel_tbu_t_x(v1, v2);
                    else if (pairID == 03) ris = Excel_tbu_t_h(v1, v2);
                    else if (pairID == 04) ris = Excel_tbu_t_vau(v1, v2);
                    else if (pairID == 06) ris = Excel_tbu_t_tr(v1, v2);
                    else if (pairID == 12) ris = Excel_tbu_ur_x(v1, v2);
                    else if (pairID == 13) ris = Excel_tbu_ur_h(v1, v2);
                    else if (pairID == 14) ris = Excel_tbu_ur_vau(v1, v2);
                    else if (pairID == 16) ris = Excel_tbu_ur_tr(v1, v2);
                    else if (pairID == 23) ris = Excel_tbu_x_h(v1, v2);
                    else if (pairID == 24) ris = Excel_tbu_x_vau(v1, v2);
                    else if (pairID == 26) ris = Excel_tbu_x_tr(v1, v2);
                    else if (pairID == 34) ris = Excel_tbu_h_vau(v1, v2);
                    else if (pairID == 36) ris = Excel_tbu_h_tr(v1, v2);
                    else if (pairID == 46) ris = Excel_tbu_vau_tr(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                case 6: // --- TARGET tr [6] ---
                    if (pairID == 01) ris = Excel_tr_t_ur(v1, v2);
                    else if (pairID == 02) ris = Excel_tr_t_x(v1, v2);
                    else if (pairID == 03) ris = Excel_tr_t_h(v1, v2);
                    else if (pairID == 04) ris = Excel_tr_t_vau(v1, v2);
                    else if (pairID == 05) ris = Excel_tr_t_tbu(v1, v2);
                    else if (pairID == 12) ris = Excel_tr_ur_x(v1, v2);
                    else if (pairID == 13) ris = Excel_tr_ur_h(v1, v2);
                    else if (pairID == 14) ris = Excel_tr_ur_vau(v1, v2);
                    else if (pairID == 15) ris = Excel_tr_ur_tbu(v1, v2);
                    else if (pairID == 23) ris = Excel_tr_x_h(v1, v2);
                    else if (pairID == 24) ris = Excel_tr_x_vau(v1, v2);
                    else if (pairID == 25) ris = Excel_tr_x_tbu(v1, v2);
                    else if (pairID == 34) ris = Excel_tr_h_vau(v1, v2);
                    else if (pairID == 35) ris = Excel_tr_h_tbu(v1, v2);
                    else if (pairID == 45) ris = Excel_tr_vau_tbu(v1, v2);
                    else return "#COMB_NON_SUPP#";
                    break;

                default:
                    return "#TARGET_ID_ERR#";
            }

            return ris;
        }
        catch (Exception ex)
        {
            return "Errore: " + ex.Message;
        }
    }

    private static int GetPropIndex(string p)
    {
        if (string.IsNullOrEmpty(p)) return -1;
        switch (p.Trim().ToLower())
        {
            case "t": return 0;
            case "ur": return 1;
            case "x": return 2;
            case "h": return 3;
            case "vau": return 4;
            case "tbu": return 5;
            case "tr": return 6;
            default: return -1;
        }
    }



    [ExcelFunction(Description = "Imposta la quota e aggiorna la pressione globale", Category = "Psicro", IsVolatile = true)]
    public static string PSICRO_SET_QUOTA(double metri)
    {
        try
        {
            // 1. Validazione dell'input
            if (metri < -430) // Quota minima terrestre (Mar Morto)
            {
                return "Errore: Quota troppo bassa (< -430m)";
            }
            if (metri > 11000) // Limite della troposfera
            {
                return "Errore: Quota troppo alta (> 11000m)";
            }

            // 2. Chiamata alla DLL C
            double pCalcolata = Excel_set_quota(metri);

            // 3. Feedback all'utente
            return $"Quota OK: {metri}m (P: {pCalcolata:F2} kPa)";
        }
        catch (Exception ex)
        {
            return "Errore critico: " + ex.Message;
        }
    }

    [ExcelFunction(Name = "PSICRO.INFO", Description = "Mostra informazioni su Psicro", Category = "Psicrometria")]
    public static string Info()
    {
        string message = "Questo piccolo Addin ti permette di calcolare una variabile psicrometrica nota altre due. \n\n Esempio =psicro.h_t_ur(26,50) \n\n Restituisce l'entalpia dell'aria umida  a 26 gradi e 50% ur \n\n UNITA' MISURA:\n";
        message = message + "\n - T/Tr/Tbu in °C";
        message = message + "\n - UR 0-100";
        message = message + "\n - X in Kgv/Kgas";
        message = message + "\n - H in kJ/Kgas";
        message = message + "\n - Vau in mc/Kgas";
        message = message + "\n\n - Pressione std 101.325KPa (per ora...)";
        message = message + "\n\n BY: Ferdinando Battillocchi 2025";
        MessageBox.Show(message, "INFO SU PsicroAddIn:");
        return message;
    }
    // Base
    //[ExcelFunction(Name = "PSICRO.PATM_Z", Category = CAT)] public static double Patm_z(double altitude) => Excel_get_patm_at_altitude(altitude);
    [ExcelFunction(Name = "PSICRO.PSAT", Category = CAT)] public static double Psat(double t) => Excel_Psat(t);
    [ExcelFunction(Name = "PSICRO.TPSAT", Category = CAT)] public static double TPsat(double p) => Excel_TPsat(p);
    [ExcelFunction(Name = "PSICRO.XSAT_T", Category = CAT)] public static double Xsat_t(double t) => Excel_xsat_t(t);

    // T - Combinazioni
    [ExcelFunction(Name = "PSICRO.T_UR_X", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)] public static double T_UR_X(
    [ExcelArgument(Description = "Umidità relativa [0-100 %]")] double ur, [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x)  => Excel_t_ur_x(ur, x);

    [ExcelFunction(Name = "PSICRO.T_UR_H", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_UR_H(
        [ExcelArgument(Description = "Umidità relativa [0-100 %]")] double ur,
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h)
        => Excel_t_ur_h(ur, h);

    [ExcelFunction(Name = "PSICRO.T_UR_VAU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_UR_VAU(
        [ExcelArgument(Description = "Umidità relativa [0-100 %]")] double ur,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_t_ur_vau(ur, vau);

    [ExcelFunction(Name = "PSICRO.T_UR_TBU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_UR_TBU(
        [ExcelArgument(Description = "Umidità relativa [0-100 %]")] double ur,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_t_ur_tbu(ur, tbu);

    [ExcelFunction(Name = "PSICRO.T_UR_TR", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_UR_TR(
        [ExcelArgument(Description = "Umidità relativa [0-100 %]")] double ur,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_t_ur_tr(ur, tr);

    [ExcelFunction(Name = "PSICRO.T_X_H", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_X_H(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h)
        => Excel_t_x_h(x, h);

    [ExcelFunction(Name = "PSICRO.T_X_VAU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_X_VAU(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_t_x_vau(x, vau);

    [ExcelFunction(Name = "PSICRO.T_X_TBU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_X_TBU(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_t_x_tbu(x, tbu);

    [ExcelFunction(Name = "PSICRO.T_X_TR", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_X_TR(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_t_x_tr(x, tr);

    [ExcelFunction(Name = "PSICRO.T_H_VAU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_H_VAU(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_t_h_vau(h, vau);

    [ExcelFunction(Name = "PSICRO.T_H_TBU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_H_TBU(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_t_h_tbu(h, tbu);

    [ExcelFunction(Name = "PSICRO.T_H_TR", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_H_TR(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_t_h_tr(h, tr);

    [ExcelFunction(Name = "PSICRO.T_VAU_TBU", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_VAU_TBU(
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_t_vau_tbu(vau, tbu);

    [ExcelFunction(Name = "PSICRO.T_VAU_TR", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_VAU_TR(
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_t_vau_tr(vau, tr);

    [ExcelFunction(Name = "PSICRO.T_TBU_TR", Description = "Calcola la temperatura a bulbo secco [°C]", Category = CAT)]
    public static double T_TBU_TR(
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_t_tbu_tr(tbu, tr);

    // UR - Combinazioni
    [ExcelFunction(Name = "PSICRO.UR_T_X", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_T_X(
    [ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,
    [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x)
    => Excel_ur_t_x(t, x);

    [ExcelFunction(Name = "PSICRO.UR_T_H", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_T_H(
        [ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h)
        => Excel_ur_t_h(t, h);

    [ExcelFunction(Name = "PSICRO.UR_T_VAU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_T_VAU(
        [ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_ur_t_vau(t, vau);

    [ExcelFunction(Name = "PSICRO.UR_T_TBU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_T_TBU(
        [ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_ur_t_tbu(t, tbu);

    [ExcelFunction(Name = "PSICRO.UR_T_TR", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_T_TR(
        [ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_ur_t_tr(t, tr);

    [ExcelFunction(Name = "PSICRO.UR_X_H", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_X_H(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h)
        => Excel_ur_x_h(x, h);

    [ExcelFunction(Name = "PSICRO.UR_X_VAU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_X_VAU(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_ur_x_vau(x, vau);

    [ExcelFunction(Name = "PSICRO.UR_X_TBU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_X_TBU(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_ur_x_tbu(x, tbu);

    [ExcelFunction(Name = "PSICRO.UR_X_TR", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_X_TR(
        [ExcelArgument(Description = "Umidità specifica [kgv/kgas]")] double x,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_ur_x_tr(x, tr);

    [ExcelFunction(Name = "PSICRO.UR_H_VAU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_H_VAU(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau)
        => Excel_ur_h_vau(h, vau);

    [ExcelFunction(Name = "PSICRO.UR_H_TBU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_H_TBU(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_ur_h_tbu(h, tbu);

    [ExcelFunction(Name = "PSICRO.UR_H_TR", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_H_TR(
        [ExcelArgument(Description = "Entalpia [kJ/kgas]")] double h,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_ur_h_tr(h, tr);

    [ExcelFunction(Name = "PSICRO.UR_VAU_TBU", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_VAU_TBU(
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau,
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu)
        => Excel_ur_vau_tbu(vau, tbu);

    [ExcelFunction(Name = "PSICRO.UR_VAU_TR", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_VAU_TR(
        [ExcelArgument(Description = "Volume specifico [m³/kgas]")] double vau,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_ur_vau_tr(vau, tr);

    [ExcelFunction(Name = "PSICRO.UR_TBU_TR", Description = "Calcola l'umidità relativa [%]", Category = CAT)]
    public static double UR_TBU_TR(
        [ExcelArgument(Description = "Temperatura di bulbo umido [°C]")] double tbu,
        [ExcelArgument(Description = "Temperatura di rugiada [°C]")] double tr)
        => Excel_ur_tbu_tr(tbu, tr);


    // X - Combinazioni

    [ExcelFunction(Name = "PSICRO.X_T_UR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_T_UR(double t, double ur) => Excel_x_t_ur(t, ur);

    [ExcelFunction(Name = "PSICRO.X_T_H", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_T_H(double t, double h) => Excel_x_t_h(t, h);

    [ExcelFunction(Name = "PSICRO.X_T_VAU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_T_VAU(double t, double vau) => Excel_x_t_vau(t, vau);

    [ExcelFunction(Name = "PSICRO.X_T_TBU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_T_TBU(double t, double tbu) => Excel_x_t_tbu(t, tbu);

    [ExcelFunction(Name = "PSICRO.X_T_TR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_T_TR(double t, double tr) => Excel_x_t_tr(t, tr);

    [ExcelFunction(Name = "PSICRO.X_UR_H", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_UR_H(double ur, double h) => Excel_x_ur_h(ur, h);

    [ExcelFunction(Name = "PSICRO.X_UR_VAU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_UR_VAU(double ur, double vau) => Excel_x_ur_vau(ur, vau);

    [ExcelFunction(Name = "PSICRO.X_UR_TBU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_UR_TBU(double ur, double tbu) => Excel_x_ur_tbu(ur, tbu);

    [ExcelFunction(Name = "PSICRO.X_UR_TR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_UR_TR(double ur, double tr) => Excel_x_ur_tr(ur, tr);

    [ExcelFunction(Name = "PSICRO.X_H_VAU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_H_VAU(double h, double vau) => Excel_x_h_vau(h, vau);

    [ExcelFunction(Name = "PSICRO.X_H_TBU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_H_TBU(double h, double tbu) => Excel_x_h_tbu(h, tbu);

    [ExcelFunction(Name = "PSICRO.X_H_TR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_H_TR(double h, double tr) => Excel_x_h_tr(h, tr);

    [ExcelFunction(Name = "PSICRO.X_VAU_TBU", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_VAU_TBU(double vau, double tbu) => Excel_x_vau_tbu(vau, tbu);

    [ExcelFunction(Name = "PSICRO.X_VAU_TR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_VAU_TR(double vau, double tr) => Excel_x_vau_tr(vau, tr);

    [ExcelFunction(Name = "PSICRO.X_TBU_TR", Description = "Calcola l'umidità specifica [kgv/kgas]", Category = CAT)]
    public static double X_TBU_TR(double tbu, double tr) => Excel_x_tbu_tr(tbu, tr);




    /*

    [ExcelFunction(Name = "PSICRO.X_T_UR", Category = CAT)] public static double X_T_UR(double t, double ur) => Excel_x_t_ur(t, ur);
    [ExcelFunction(Name = "PSICRO.X_T_H", Category = CAT)] public static double X_T_H(double t, double h) => Excel_x_t_h(t, h);
    [ExcelFunction(Name = "PSICRO.X_T_VAU", Category = CAT)] public static double X_T_VAU(double t, double v) => Excel_x_t_vau(t, v);
    [ExcelFunction(Name = "PSICRO.X_T_TBU", Category = CAT)] public static double X_T_TBU(double t, double tb) => Excel_x_t_tbu(t, tb);
    [ExcelFunction(Name = "PSICRO.X_T_TR", Category = CAT)] public static double X_T_TR(double t, double tr) => Excel_x_t_tr(t, tr);
    [ExcelFunction(Name = "PSICRO.X_UR_H", Category = CAT)] public static double X_UR_H(double ur, double h) => Excel_x_ur_h(ur, h);
    [ExcelFunction(Name = "PSICRO.X_UR_VAU", Category = CAT)] public static double X_UR_VAU(double ur, double v) => Excel_x_ur_vau(ur, v);
    [ExcelFunction(Name = "PSICRO.X_UR_TBU", Category = CAT)] public static double X_UR_TBU(double ur, double t) => Excel_x_ur_tbu(ur, t);
    [ExcelFunction(Name = "PSICRO.X_UR_TR", Category = CAT)] public static double X_UR_TR(double ur, double tr) => Excel_x_ur_tr(ur, tr);
    [ExcelFunction(Name = "PSICRO.X_H_VAU", Category = CAT)] public static double X_H_VAU(double h, double v) => Excel_x_h_vau(h, v);
    [ExcelFunction(Name = "PSICRO.X_H_TBU", Category = CAT)] public static double X_H_TBU(double h, double t) => Excel_x_h_tbu(h, t);
    [ExcelFunction(Name = "PSICRO.X_H_TR", Category = CAT)] public static double X_H_TR(double h, double tr) => Excel_x_h_tr(h, tr);
    [ExcelFunction(Name = "PSICRO.X_VAU_TBU", Category = CAT)] public static double X_VAU_TBU(double v, double t) => Excel_x_vau_tbu(v, t);
    [ExcelFunction(Name = "PSICRO.X_VAU_TR", Category = CAT)] public static double X_VAU_TR(double v, double tr) => Excel_x_vau_tr(v, tr);
    [ExcelFunction(Name = "PSICRO.X_TBU_TR", Category = CAT)] public static double X_TBU_TR(double t, double tr) => Excel_x_tbu_tr(t, tr);
    */
    // H - Combinazioni
    [ExcelFunction(Name = "PSICRO.H_T_UR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_T_UR(double t, double ur) => Excel_h_t_ur(t, ur);

    [ExcelFunction(Name = "PSICRO.H_T_X", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_T_X(double t, double x) => Excel_h_t_x(t, x);

    [ExcelFunction(Name = "PSICRO.H_T_VAU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_T_VAU(double t, double vau) => Excel_h_t_vau(t, vau);

    [ExcelFunction(Name = "PSICRO.H_T_TBU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_T_TBU(double t, double tbu) => Excel_h_t_tbu(t, tbu);

    [ExcelFunction(Name = "PSICRO.H_T_TR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_T_TR(double t, double tr) => Excel_h_t_tr(t, tr);

    [ExcelFunction(Name = "PSICRO.H_UR_X", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_UR_X(double ur, double x) => Excel_h_ur_x(ur, x);

    [ExcelFunction(Name = "PSICRO.H_UR_VAU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_UR_VAU(double ur, double vau) => Excel_h_ur_vau(ur, vau);

    [ExcelFunction(Name = "PSICRO.H_UR_TBU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_UR_TBU(double ur, double tbu) => Excel_h_ur_tbu(ur, tbu);

    [ExcelFunction(Name = "PSICRO.H_UR_TR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_UR_TR(double ur, double tr) => Excel_h_ur_tr(ur, tr);

    [ExcelFunction(Name = "PSICRO.H_X_VAU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_X_VAU(double x, double vau) => Excel_h_x_vau(x, vau);

    [ExcelFunction(Name = "PSICRO.H_X_TBU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_X_TBU(double x, double tbu) => Excel_h_x_tbu(x, tbu);

    [ExcelFunction(Name = "PSICRO.H_X_TR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_X_TR(double x, double tr) => Excel_h_x_tr(x, tr);

    [ExcelFunction(Name = "PSICRO.H_VAU_TBU", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_VAU_TBU(double vau, double tbu) => Excel_h_vau_tbu(vau, tbu);

    [ExcelFunction(Name = "PSICRO.H_VAU_TR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_VAU_TR(double vau, double tr) => Excel_h_vau_tr(vau, tr);

    [ExcelFunction(Name = "PSICRO.H_TBU_TR", Description = "Calcola l'entalpia dell'aria umida [kJ/kgas]", Category = CAT)]
    public static double H_TBU_TR(double tbu, double tr) => Excel_h_tbu_tr(tbu, tr);

    /*
    [ExcelFunction(Name = "PSICRO.H_T_UR", Description = "Calcola l'entalpia dell'aria umida [kJ/kg]", Category = "Psicrometria")]
    public static double H_T_UR([ExcelArgument(Description = "Temperatura a bulbo secco [°C]")] double t,[ExcelArgument(Description = "Umidità relativa [0-100]")] double ur){return Excel_h_t_ur(t, ur);}

    //[ExcelFunction(Name = "PSICRO.H_T_UR", Category = CAT)] public static double H_T_UR(double t, double ur) => Excel_h_t_ur(t, ur);
    [ExcelFunction(Name = "PSICRO.H_T_X", Category = CAT)] public static double H_T_X(double t, double x) => Excel_h_t_x(t, x);
    [ExcelFunction(Name = "PSICRO.H_T_VAU", Category = CAT)] public static double H_T_VAU(double t, double v) => Excel_h_t_vau(t, v);
    [ExcelFunction(Name = "PSICRO.H_T_TBU", Category = CAT)] public static double H_T_TBU(double t, double tb) => Excel_h_t_tbu(t, tb);
    [ExcelFunction(Name = "PSICRO.H_T_TR", Category = CAT)] public static double H_T_TR(double t, double tr) => Excel_h_t_tr(t, tr);
    [ExcelFunction(Name = "PSICRO.H_UR_X", Category = CAT)] public static double H_UR_X(double ur, double x) => Excel_h_ur_x(ur, x);
    [ExcelFunction(Name = "PSICRO.H_UR_VAU", Category = CAT)] public static double H_UR_VAU(double ur, double v) => Excel_h_ur_vau(ur, v);
    [ExcelFunction(Name = "PSICRO.H_UR_TBU", Category = CAT)] public static double H_UR_TBU(double ur, double t) => Excel_h_ur_tbu(ur, t);
    [ExcelFunction(Name = "PSICRO.H_UR_TR", Category = CAT)] public static double H_UR_TR(double ur, double tr) => Excel_h_ur_tr(ur, tr);
    [ExcelFunction(Name = "PSICRO.H_X_VAU", Category = CAT)] public static double H_X_VAU(double x, double v) => Excel_h_x_vau(x, v);
    [ExcelFunction(Name = "PSICRO.H_X_TBU", Category = CAT)] public static double H_X_TBU(double x, double t) => Excel_h_x_tbu(x, t);
    [ExcelFunction(Name = "PSICRO.H_X_TR", Category = CAT)] public static double H_X_TR(double x, double tr) => Excel_h_x_tr(x, tr);
    [ExcelFunction(Name = "PSICRO.H_VAU_TBU", Category = CAT)] public static double H_VAU_TBU(double v, double t) => Excel_h_vau_tbu(v, t);
    [ExcelFunction(Name = "PSICRO.H_VAU_TR", Category = CAT)] public static double H_VAU_TR(double v, double tr) => Excel_h_vau_tr(v, tr);
    [ExcelFunction(Name = "PSICRO.H_TBU_TR", Category = CAT)] public static double H_TBU_TR(double t, double tr) => Excel_h_tbu_tr(t, tr);
    */
    // VAU - Combinazioni

    [ExcelFunction(Name = "PSICRO.VAU_T_UR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_T_UR(double t, double ur) => Excel_vau_t_ur(t, ur);

    [ExcelFunction(Name = "PSICRO.VAU_T_X", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_T_X(double t, double x) => Excel_vau_t_x(t, x);

    [ExcelFunction(Name = "PSICRO.VAU_T_H", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_T_H(double t, double h) => Excel_vau_t_h(t, h);

    [ExcelFunction(Name = "PSICRO.VAU_T_TBU", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_T_TBU(double t, double tbu) => Excel_vau_t_tbu(t, tbu);

    [ExcelFunction(Name = "PSICRO.VAU_T_TR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_T_TR(double t, double tr) => Excel_vau_t_tr(t, tr);

    [ExcelFunction(Name = "PSICRO.VAU_UR_X", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_UR_X(double ur, double x) => Excel_vau_ur_x(ur, x);

    [ExcelFunction(Name = "PSICRO.VAU_UR_H", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_UR_H(double ur, double h) => Excel_vau_ur_h(ur, h);

    [ExcelFunction(Name = "PSICRO.VAU_UR_TBU", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_UR_TBU(double ur, double tbu) => Excel_vau_ur_tbu(ur, tbu);

    [ExcelFunction(Name = "PSICRO.VAU_UR_TR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_UR_TR(double ur, double tr) => Excel_vau_ur_tr(ur, tr);

    [ExcelFunction(Name = "PSICRO.VAU_X_H", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_X_H(double x, double h) => Excel_vau_x_h(x, h);

    [ExcelFunction(Name = "PSICRO.VAU_X_TBU", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_X_TBU(double x, double tbu) => Excel_vau_x_tbu(x, tbu);

    [ExcelFunction(Name = "PSICRO.VAU_X_TR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_X_TR(double x, double tr) => Excel_vau_x_tr(x, tr);

    [ExcelFunction(Name = "PSICRO.VAU_H_TBU", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_H_TBU(double h, double tbu) => Excel_vau_h_tbu(h, tbu);

    [ExcelFunction(Name = "PSICRO.VAU_H_TR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_H_TR(double h, double tr) => Excel_vau_h_tr(h, tr);

    [ExcelFunction(Name = "PSICRO.VAU_TBU_TR", Description = "Calcola il volume specifico [m³/kgas]", Category = CAT)]
    public static double VAU_TBU_TR(double tbu, double tr) => Excel_vau_tbu_tr(tbu, tr);

    /*
    [ExcelFunction(Name = "PSICRO.VAU_T_UR", Category = CAT)] public static double VAU_T_UR(double t, double ur) => Excel_vau_t_ur(t, ur);
    [ExcelFunction(Name = "PSICRO.VAU_T_X", Category = CAT)] public static double VAU_T_X(double t, double x) => Excel_vau_t_x(t, x);
    [ExcelFunction(Name = "PSICRO.VAU_T_H", Category = CAT)] public static double VAU_T_H(double t, double h) => Excel_vau_t_h(t, h);
    [ExcelFunction(Name = "PSICRO.VAU_T_TBU", Category = CAT)] public static double VAU_T_TBU(double t, double tb) => Excel_vau_t_tbu(t, tb);
    [ExcelFunction(Name = "PSICRO.VAU_T_TR", Category = CAT)] public static double VAU_T_TR(double t, double tr) => Excel_vau_t_tr(t, tr);
    [ExcelFunction(Name = "PSICRO.VAU_UR_X", Category = CAT)] public static double VAU_UR_X(double ur, double x) => Excel_vau_ur_x(ur, x);
    [ExcelFunction(Name = "PSICRO.VAU_UR_H", Category = CAT)] public static double VAU_UR_H(double ur, double h) => Excel_vau_ur_h(ur, h);
    [ExcelFunction(Name = "PSICRO.VAU_UR_TBU", Category = CAT)] public static double VAU_UR_TBU(double ur, double t) => Excel_vau_ur_tbu(ur, t);
    [ExcelFunction(Name = "PSICRO.VAU_UR_TR", Category = CAT)] public static double VAU_UR_TR(double ur, double tr) => Excel_vau_ur_tr(ur, tr);
    [ExcelFunction(Name = "PSICRO.VAU_X_H", Category = CAT)] public static double VAU_X_H(double x, double h) => Excel_vau_x_h(x, h);
    [ExcelFunction(Name = "PSICRO.VAU_X_TBU", Category = CAT)] public static double VAU_X_TBU(double x, double t) => Excel_vau_x_tbu(x, t);
    [ExcelFunction(Name = "PSICRO.VAU_X_TR", Category = CAT)] public static double VAU_X_TR(double x, double tr) => Excel_vau_x_tr(x, tr);
    [ExcelFunction(Name = "PSICRO.VAU_H_TBU", Category = CAT)] public static double VAU_H_TBU(double h, double t) => Excel_vau_h_tbu(h, t);
    [ExcelFunction(Name = "PSICRO.VAU_H_TR", Category = CAT)] public static double VAU_H_TR(double h, double tr) => Excel_vau_h_tr(h, tr);
    [ExcelFunction(Name = "PSICRO.VAU_TBU_TR", Category = CAT)] public static double VAU_TBU_TR(double t, double tr) => Excel_vau_tbu_tr(t, tr);
    */
    // TBU - Combinazioni
    [ExcelFunction(Name = "PSICRO.TBU_T_UR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_T_UR(double t, double ur) => Excel_tbu_t_ur(t, ur);

    [ExcelFunction(Name = "PSICRO.TBU_T_X", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_T_X(double t, double x) => Excel_tbu_t_x(t, x);

    [ExcelFunction(Name = "PSICRO.TBU_T_H", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_T_H(double t, double h) => Excel_tbu_t_h(t, h);

    [ExcelFunction(Name = "PSICRO.TBU_T_VAU", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_T_VAU(double t, double vau) => Excel_tbu_t_vau(t, vau);

    [ExcelFunction(Name = "PSICRO.TBU_T_TR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_T_TR(double t, double tr) => Excel_tbu_t_tr(t, tr);

    [ExcelFunction(Name = "PSICRO.TBU_UR_X", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_UR_X(double ur, double x) => Excel_tbu_ur_x(ur, x);

    [ExcelFunction(Name = "PSICRO.TBU_UR_H", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_UR_H(double ur, double h) => Excel_tbu_ur_h(ur, h);

    [ExcelFunction(Name = "PSICRO.TBU_UR_VAU", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_UR_VAU(double ur, double vau) => Excel_tbu_ur_vau(ur, vau);

    [ExcelFunction(Name = "PSICRO.TBU_UR_TR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_UR_TR(double ur, double tr) => Excel_tbu_ur_tr(ur, tr);

    [ExcelFunction(Name = "PSICRO.TBU_X_H", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_X_H(double x, double h) => Excel_tbu_x_h(x, h);

    [ExcelFunction(Name = "PSICRO.TBU_X_VAU", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_X_VAU(double x, double vau) => Excel_tbu_x_vau(x, vau);

    [ExcelFunction(Name = "PSICRO.TBU_X_TR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_X_TR(double x, double tr) => Excel_tbu_x_tr(x, tr);

    [ExcelFunction(Name = "PSICRO.TBU_H_VAU", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_H_VAU(double h, double vau) => Excel_tbu_h_vau(h, vau);

    [ExcelFunction(Name = "PSICRO.TBU_H_TR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_H_TR(double h, double tr) => Excel_tbu_h_tr(h, tr);

    [ExcelFunction(Name = "PSICRO.TBU_VAU_TR", Description = "Calcola la temperatura di bulbo umido [°C]", Category = CAT)]
    public static double TBU_VAU_TR(double vau, double tr) => Excel_tbu_vau_tr(vau, tr);

    /*
    [ExcelFunction(Name = "PSICRO.TBU_T_UR", Category = CAT)] public static double TBU_T_UR(double t, double ur) => Excel_tbu_t_ur(t, ur);
    [ExcelFunction(Name = "PSICRO.TBU_T_X", Category = CAT)] public static double TBU_T_X(double t, double x) => Excel_tbu_t_x(t, x);
    [ExcelFunction(Name = "PSICRO.TBU_T_H", Category = CAT)] public static double TBU_T_H(double t, double h) => Excel_tbu_t_h(t, h);
    [ExcelFunction(Name = "PSICRO.TBU_T_VAU", Category = CAT)] public static double TBU_T_VAU(double t, double v) => Excel_tbu_t_vau(t, v);
    [ExcelFunction(Name = "PSICRO.TBU_T_TR", Category = CAT)] public static double TBU_T_TR(double t, double tr) => Excel_tbu_t_tr(t, tr);
    [ExcelFunction(Name = "PSICRO.TBU_UR_X", Category = CAT)] public static double TBU_UR_X(double ur, double x) => Excel_tbu_ur_x(ur, x);
    [ExcelFunction(Name = "PSICRO.TBU_UR_H", Category = CAT)] public static double TBU_UR_H(double ur, double h) => Excel_tbu_ur_h(ur, h);
    [ExcelFunction(Name = "PSICRO.TBU_UR_VAU", Category = CAT)] public static double TBU_UR_VAU(double ur, double v) => Excel_tbu_ur_vau(ur, v);
    [ExcelFunction(Name = "PSICRO.TBU_UR_TR", Category = CAT)] public static double TBU_UR_TR(double ur, double tr) => Excel_tbu_ur_tr(ur, tr);
    [ExcelFunction(Name = "PSICRO.TBU_X_H", Category = CAT)] public static double TBU_X_H(double x, double h) => Excel_tbu_x_h(x, h);
    [ExcelFunction(Name = "PSICRO.TBU_X_VAU", Category = CAT)] public static double TBU_X_VAU(double x, double v) => Excel_tbu_x_vau(x, v);
    [ExcelFunction(Name = "PSICRO.TBU_X_TR", Category = CAT)] public static double TBU_X_TR(double x, double tr) => Excel_tbu_x_tr(x, tr);
    [ExcelFunction(Name = "PSICRO.TBU_H_VAU", Category = CAT)] public static double TBU_H_VAU(double h, double v) => Excel_tbu_h_vau(h, v);
    [ExcelFunction(Name = "PSICRO.TBU_H_TR", Category = CAT)] public static double TBU_H_TR(double h, double tr) => Excel_tbu_h_tr(h, tr);
    [ExcelFunction(Name = "PSICRO.TBU_VAU_TR", Category = CAT)] public static double TBU_VAU_TR(double v, double tr) => Excel_tbu_vau_tr(v, tr);
    */
    // TR - Combinazioni
    /*
    [ExcelFunction(Name = "PSICRO.TR_T_UR", Category = CAT)] public static double TR_T_UR(double t, double ur) => Excel_tr_t_ur(t, ur);
    [ExcelFunction(Name = "PSICRO.TR_T_X", Category = CAT)] public static double TR_T_X(double t, double x) => Excel_tr_t_x(t, x);
    [ExcelFunction(Name = "PSICRO.TR_T_H", Category = CAT)] public static double TR_T_H(double t, double h) => Excel_tr_t_h(t, h);
    [ExcelFunction(Name = "PSICRO.TR_T_VAU", Category = CAT)] public static double TR_T_VAU(double t, double v) => Excel_tr_t_vau(t, v);
    [ExcelFunction(Name = "PSICRO.TR_T_TBU", Category = CAT)] public static double TR_T_TBU(double t, double tb) => Excel_tr_t_tbu(t, tb);
    [ExcelFunction(Name = "PSICRO.TR_UR_X", Category = CAT)] public static double TR_UR_X(double ur, double x) => Excel_tr_ur_x(ur, x);
    [ExcelFunction(Name = "PSICRO.TR_UR_H", Category = CAT)] public static double TR_UR_H(double ur, double h) => Excel_tr_ur_h(ur, h);
    [ExcelFunction(Name = "PSICRO.TR_UR_VAU", Category = CAT)] public static double TR_UR_VAU(double ur, double v) => Excel_tr_ur_vau(ur, v);
    [ExcelFunction(Name = "PSICRO.TR_UR_TBU", Category = CAT)] public static double TR_UR_TBU(double ur, double t) => Excel_tr_ur_tbu(ur, t);
    [ExcelFunction(Name = "PSICRO.TR_X_H", Category = CAT)] public static double TR_X_H(double x, double h) => Excel_tr_x_h(x, h);
    [ExcelFunction(Name = "PSICRO.TR_X_VAU", Category = CAT)] public static double TR_X_VAU(double x, double v) => Excel_tr_x_vau(x, v);
    [ExcelFunction(Name = "PSICRO.TR_X_TBU", Category = CAT)] public static double TR_X_TBU(double x, double t) => Excel_tr_x_tbu(x, t);
    [ExcelFunction(Name = "PSICRO.TR_H_VAU", Category = CAT)] public static double TR_H_VAU(double h, double v) => Excel_tr_h_vau(h, v);
    [ExcelFunction(Name = "PSICRO.TR_H_TBU", Category = CAT)] public static double TR_H_TBU(double h, double t) => Excel_tr_h_tbu(h, t);
    [ExcelFunction(Name = "PSICRO.TR_VAU_TBU", Category = CAT)] public static double TR_VAU_TBU(double v, double t) => Excel_tr_vau_tbu(v, t);
    */
    [ExcelFunction(Name = "PSICRO.TR_T_UR", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_T_UR(double t, double ur) => Excel_tr_t_ur(t, ur);

    [ExcelFunction(Name = "PSICRO.TR_T_X", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_T_X(double t, double x) => Excel_tr_t_x(t, x);

    [ExcelFunction(Name = "PSICRO.TR_T_H", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_T_H(double t, double h) => Excel_tr_t_h(t, h);

    [ExcelFunction(Name = "PSICRO.TR_T_VAU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_T_VAU(double t, double vau) => Excel_tr_t_vau(t, vau);

    [ExcelFunction(Name = "PSICRO.TR_T_TBU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_T_TBU(double t, double tbu) => Excel_tr_t_tbu(t, tbu);

    [ExcelFunction(Name = "PSICRO.TR_UR_X", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_UR_X(double ur, double x) => Excel_tr_ur_x(ur, x);

    [ExcelFunction(Name = "PSICRO.TR_UR_H", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_UR_H(double ur, double h) => Excel_tr_ur_h(ur, h);

    [ExcelFunction(Name = "PSICRO.TR_UR_VAU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_UR_VAU(double ur, double vau) => Excel_tr_ur_vau(ur, vau);

    [ExcelFunction(Name = "PSICRO.TR_UR_TBU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_UR_TBU(double ur, double tbu) => Excel_tr_ur_tbu(ur, tbu);

    [ExcelFunction(Name = "PSICRO.TR_X_H", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_X_H(double x, double h) => Excel_tr_x_h(x, h);

    [ExcelFunction(Name = "PSICRO.TR_X_VAU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_X_VAU(double x, double vau) => Excel_tr_x_vau(x, vau);

    [ExcelFunction(Name = "PSICRO.TR_X_TBU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_X_TBU(double x, double tbu) => Excel_tr_x_tbu(x, tbu);

    [ExcelFunction(Name = "PSICRO.TR_H_VAU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_H_VAU(double h, double vau) => Excel_tr_h_vau(h, vau);

    [ExcelFunction(Name = "PSICRO.TR_H_TBU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_H_TBU(double h, double tbu) => Excel_tr_h_tbu(h, tbu);

    [ExcelFunction(Name = "PSICRO.TR_VAU_TBU", Description = "Calcola la temperatura di rugiada [°C]", Category = CAT)]
    public static double TR_VAU_TBU(double vau, double tbu) => Excel_tr_vau_tbu(vau, tbu);

    #endregion

}