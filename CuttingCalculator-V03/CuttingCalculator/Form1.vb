Option Strict Off
Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing


Public Class Form1
    Public spec_snijkracht_0, diameter, snijsnelheid, toerental, aanzetpertand, aanzet, aantaltanden, aanzetperomw As Double
    Public snedediepte, snedebreedte, spiraalhoek, spaanhoek, gem_spaandikte, verspaand_vol, spaandoorsnede, werkelijke_spaandikte, werkelijke_snedediepte As Double
    Public tanden_eff, spec_snijkracht_g, snijkracht_r, snijkracht_a, snijkracht_t, vermogen, koppel As Double
    Public snedebreedtep, snedebreedtem, BrutoVermogen, snedehoek, snedebreedteprocent, mc As Double
    Public N2V As Boolean
    Dim memoryImage As Bitmap

    Private WithEvents printDocument1 As New PrintDocument

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Clear()
    End Sub

    Private Sub Clear()
        If TabControl1.SelectedIndex = 0 Then
            d.Text = 0
            V.Text = 0
            Z.Text = 4
            ft.Text = 0
            gama.Text = 0
            alfa.Text = 0
            Ap.Text = 0
            Ae.Text = ""
            Aem.Text = ""
            f.Text = ""
            n.Text = ""
            hm.Text = ""
            Q.Text = ""
            A.Text = ""
            Ze.Text = ""
            Kcg.Text = ""
            Fc.Text = ""
            Fa.Text = ""
            Fr.Text = ""
            Pn.Text = ""
            PB.Text = ""
            M.Text = ""
        End If
        If TabControl1.SelectedIndex = 1 Then
            turn_A.Text = ""
            turn_Ap.Text = 0
            turn_Apw.Text = ""
            turn_d.Text = 0
            turn_Fa.Text = ""
            turn_Fc.Text = ""
            turn_Fn.Text = ""
            turn_Fnw.Text = ""
            turn_Fr.Text = ""
            turn_gama.Text = 0
            turn_kappa.Text = 0
            turn_Kcg.Text = ""
            turn_n.Text = ""
            turn_Pb.Text = ""
            turn_Pn.Text = ""
            turn_Q.Text = ""
            turn_Tq.Text = ""
            turn_V.Text = 0
        End If
        If TabControl1.SelectedIndex = 2 Then
            dril_A.Text = ""
            dril_d.Text = 0
            dril_F.Text = ""
            dril_fa.Text = ""
            dril_fc.Text = ""
            dril_fn.Text = 0
            dril_Fnw.Text = ""
            dril_fr.Text = ""
            dril_kcg.Text = ""
            dril_m.Text = ""
            dril_n.Text = ""
            dril_pb.Text = ""
            dril_pn.Text = ""
            dril_Q.Text = ""
            dril_v.Text = 0
            dril_snedehoek.Text = ""
        End If

    End Sub
    Private Sub Calculate()

        If TabControl1.SelectedIndex = 0 Then Calculate_Mill()
        If TabControl1.SelectedIndex = 1 Then Calculate_Turn()
        If TabControl1.SelectedIndex = 2 Then Calculate_Drill()

    End Sub

    Private Sub Switch_N_V()
        If N2V = False Then
            'Clear()
            n.BackColor = d.BackColor
            f.BackColor = d.BackColor
            V.BackColor = Q.BackColor
            ft.BackColor = Q.BackColor
            N2V = True
        Else
            'Clear()
            n.BackColor = Q.BackColor
            f.BackColor = Q.BackColor
            V.BackColor = d.BackColor
            ft.BackColor = d.BackColor
            N2V = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Calculate()
    End Sub

    Private Sub Calculate_Mill()

        On Error Resume Next
        spec_snijkracht_0 = CDbl(Replace(Kc0.Text, ".", ","))
        diameter = CDbl(Replace(d.Text, ".", ","))
        aantaltanden = CInt(Replace(Z.Text, ".", ","))
        spiraalhoek = CInt(Replace(alfa.Text, ".", ","))
        spaanhoek = CInt(Replace(gama.Text, ".", ","))
        snijsnelheid = CDbl(Replace(V.Text, ".", ","))
        aanzetpertand = CDbl(Replace(ft.Text, ".", ","))
        toerental = CDbl(Replace(n.Text, ".", ","))
        aanzet = CDbl(Replace(f.Text, ".", ","))
        snedediepte = CDbl(Replace(Ap.Text, ".", ","))
        snedebreedtep = CDbl(Replace(Ae.Text, ".", ",")) / 100 * diameter
        snedebreedtem = CDbl(Replace(Aem.Text, ".", ","))
        snedehoek = CDbl(Replace(kappa.Text, ".", ","))
        mc = CDbl(Replace(mcfac.Text, ".", ","))

        If diameter = 0 Then diameter = 1 'tbv niet delen door 0
        If Ae.Text = "" Then snedebreedte = snedebreedtem
        If Aem.Text = "" Then snedebreedte = snedebreedtep
        snedebreedteprocent = snedebreedte / diameter
        If snedebreedteprocent = 0 Then snedebreedteprocent = 1 'tbv niet delen door 0


        If N2V Then
            snijsnelheid = Calc_V(toerental, diameter)
            aanzetpertand = Calc_fz(toerental, aanzet, aantaltanden)
        Else
            toerental = Calc_n(snijsnelheid, diameter)
            aanzet = Calc_f(toerental, aanzetpertand, aantaltanden)
        End If

        gem_spaandikte = Calc_Hm(aanzetpertand, snedebreedteprocent, snedehoek, diameter)
        verspaand_vol = Calc_Q(snedediepte, snedebreedte, aanzet)
        spaandoorsnede = Calc_A(snedediepte, gem_spaandikte)
        tanden_eff = Calc_Ze(aantaltanden, snedebreedteprocent)

        Dim GemiddeldeSpaanDikte As Double = gem_spaandikte
        If GemiddeldeSpaanDikte = 0 Then GemiddeldeSpaanDikte = 1 'tbv niet delen door 0
        spec_snijkracht_g = Calc_Kcg(spec_snijkracht_0, spaanhoek, GemiddeldeSpaanDikte, mc)

        snijkracht_t = Calc_Ft(spec_snijkracht_g, spaandoorsnede, tanden_eff)
        snijkracht_r = Calc_Fr(snijkracht_t, spiraalhoek)
        snijkracht_a = Calc_Fa(snijkracht_t, spiraalhoek)
        'vermogen = Calc_Pn(snijkracht_t, snijsnelheid)
        vermogen = Calc_Pnet(snedediepte, snedebreedte, aanzet, snijkracht_t)
        BrutoVermogen = vermogen / 0.95
        koppel = Calc_M(snijkracht_t, diameter)

        If vermogen < 1 Then
            vermogen = vermogen * 1000
            KWlabel.Text = "W"
        Else
            KWlabel.Text = "Kw"
        End If

        If N2V Then
            V.Text = Math.Round(snijsnelheid, 0)
            ft.Text = Math.Round(aanzetpertand, 3)
        Else
            n.Text = Math.Round(toerental, 0)
            f.Text = Math.Round(aanzet, 0)
        End If
        hm.Text = Math.Round(gem_spaandikte, 3)
        Q.Text = Math.Round(verspaand_vol, 1)
        A.Text = Math.Round(spaandoorsnede, 2)
        Ze.Text = Math.Round(tanden_eff, 2)
        Kcg.Text = Math.Round(spec_snijkracht_g, 0)
        Fc.Text = Math.Round(snijkracht_t, 0)
        Fa.Text = Math.Round(snijkracht_a, 0)
        Fr.Text = Math.Round(snijkracht_r, 0)
        Pn.Text = Math.Round(vermogen, 3)
        PB.Text = Math.Round(BrutoVermogen, 2)
        M.Text = Math.Round(koppel, 1)

    End Sub

    Private Sub Calculate_Turn()

        On Error Resume Next
        spec_snijkracht_0 = CDbl(Replace(turn_Kc.Text, ".", ","))
        diameter = CDbl(Replace(turn_d.Text, ".", ","))
        spaanhoek = CInt(Replace(turn_gama.Text, ".", ","))
        snijsnelheid = CDbl(Replace(turn_V.Text, ".", ","))
        aanzetperomw = CDbl(Replace(turn_Fn.Text, ".", ","))
        snedediepte = CDbl(Replace(turn_Ap.Text, ".", ","))
        snedehoek = CDbl(Replace(turn_kappa.Text, ".", ","))
        mc = CDbl(Replace(turn_mcfac.Text, ".", ","))

        If diameter = 0 Then diameter = 1 'tbv niet delen door 0

        toerental = Calc_n(snijsnelheid, diameter)
        'aanzet = Calc_f(toerental, aanzetpertand, aantaltanden)
        werkelijke_spaandikte = Calc_T_Fnw(aanzetperomw, snedehoek)
        werkelijke_snedediepte = Calc_T_Apw(snedediepte, snedehoek)
        verspaand_vol = Calc_T_Q(snedediepte, aanzetperomw, snijsnelheid)
        spaandoorsnede = Calc_T_A(snedediepte, aanzetperomw)

        spec_snijkracht_g = Calc_Kcg(spec_snijkracht_0, spaanhoek, werkelijke_spaandikte, mc)

        snijkracht_t = Calc_T_Ft(spec_snijkracht_g, spaandoorsnede)
        snijkracht_r = Calc_T_Fr(snijkracht_t, snedehoek)
        snijkracht_a = Calc_T_Fa(snijkracht_t, snedehoek)
        vermogen = Calc_T_Pnet(spec_snijkracht_g, snijsnelheid, spaandoorsnede)
        BrutoVermogen = vermogen / 0.75
        koppel = Calc_T_M(snijkracht_t, diameter)

        If vermogen < 1 Then
            vermogen = vermogen * 1000
            T_KW_Label.Text = "W"
        Else
            T_KW_Label.Text = "Kw"
        End If

        turn_n.Text = Math.Round(toerental, 0)
        turn_Apw.Text = Math.Round(werkelijke_snedediepte, 2)
        turn_Fnw.Text = Math.Round(werkelijke_spaandikte, 3)
        turn_Q.Text = Math.Round(verspaand_vol, 1)
        turn_A.Text = Math.Round(spaandoorsnede, 2)
        turn_Kcg.Text = Math.Round(spec_snijkracht_g, 0)
        turn_Fc.Text = Math.Round(snijkracht_t, 0)
        turn_Fa.Text = Math.Round(snijkracht_a, 0)
        turn_Fr.Text = Math.Round(snijkracht_r, 0)
        turn_Pn.Text = Math.Round(vermogen, 3)
        turn_Pb.Text = Math.Round(BrutoVermogen, 2)
        turn_Tq.Text = Math.Round(koppel, 1)

    End Sub

    Private Sub Calculate_Drill()
        On Error Resume Next

        spec_snijkracht_0 = CDbl(Replace(dril_kc.Text, ".", ","))
        diameter = CDbl(Replace(dril_d.Text, ".", ","))
        aantaltanden = CInt(Replace(dril_z.Text, ".", ","))
        spaanhoek = CInt(Replace(dril_gama.Text, ".", ","))
        snijsnelheid = CDbl(Replace(dril_v.Text, ".", ","))
        aanzetperomw = CDbl(Replace(dril_fn.Text, ".", ","))
        snedehoek = (180 - CDbl(Replace(dril_kappa.Text, ".", ","))) / 2
        mc = CDbl(Replace(dril_mcfac.Text, ".", ","))

        If diameter = 0 Then diameter = 1 'tbv niet delen door 0

        toerental = Calc_n(snijsnelheid, diameter)
        aanzet = Calc_D_f(toerental, aanzetperomw)
        werkelijke_spaandikte = Calc_T_Fnw(aanzetperomw, snedehoek)
        verspaand_vol = Calc_D_Q(diameter, aanzet)
        spaandoorsnede = Calc_D_A(diameter, aanzetperomw)

        spec_snijkracht_g = Calc_Kcg(spec_snijkracht_0, spaanhoek, werkelijke_spaandikte, mc)

        snijkracht_t = Calc_Ft(spec_snijkracht_g, spaandoorsnede, aantaltanden)
        snijkracht_r = Calc_Fr(snijkracht_t, snedehoek)
        snijkracht_a = Calc_Fa(snijkracht_t, snedehoek)

        vermogen = Calc_D_Pnet(aanzetperomw, snijsnelheid, diameter, spec_snijkracht_0)
        BrutoVermogen = vermogen / 0.95
        koppel = Calc_M(snijkracht_t, diameter)

        If vermogen < 1 Then
            vermogen = vermogen * 1000
            dril_KW_label.Text = "W"
        Else
            dril_KW_label.Text = "Kw"
        End If

        dril_n.Text = Math.Round(toerental, 0)
        dril_F.Text = Math.Round(aanzet, 0)
        dril_Fnw.Text = Math.Round(werkelijke_spaandikte, 3)
        dril_Q.Text = Math.Round(verspaand_vol, 1)
        dril_A.Text = Math.Round(spaandoorsnede, 2)
        dril_kcg.Text = Math.Round(spec_snijkracht_g, 0)

        If aanzetperomw = 0 Then Exit Sub

        dril_fc.Text = Math.Round(snijkracht_t, 0)
        dril_fa.Text = Math.Round(snijkracht_a, 0)
        dril_fr.Text = Math.Round(snijkracht_r, 0)
        dril_pn.Text = Math.Round(vermogen, 3)
        dril_pb.Text = Math.Round(BrutoVermogen, 2)
        dril_m.Text = Math.Round(koppel, 1)
        dril_m.Text = Math.Round(koppel, 1)
        dril_snedehoek.Text = Math.Round(snedehoek, 1)

    End Sub


    '################################################## Rekenformules ###########################################################

    'Toerental ### MILL & TURN & Drill ###
    Private Function Calc_n(_V_ As Double, _D_ As Double) As Double
        Return _V_ * 1000 / (Math.PI * _D_)
    End Function

    'Snijsnelheid ### MILL & TURN & Drill ###
    Private Function Calc_V(_n_ As Double, _D_ As Double) As Double
        Return _n_ * Math.PI * (_D_ / 1000)
    End Function

    'Voeding  ### MILL ###
    Private Function Calc_f(_n_ As Double, _ft_ As Double, _t_ As Double) As Double
        Return _n_ * _ft_ * _t_
    End Function

    'Voeding per tand  ### MILL ###
    Private Function Calc_fz(_n_ As Double, _f_ As Double, _t_ As Double) As Double
        Return _f_ / (_n_ * _t_)
    End Function

    'Voeding  ### DRILL ###
    Private Function Calc_D_f(_n_ As Double, _fn_ As Double) As Double
        Return _n_ * _fn_
    End Function

    'Gemiddelde spaandikte  ### MILL ###
    Private Function Calc_Hm(_ft_ As Double, _Ae_ As Double, _Kap_ As Double, _D_ As Double) As Double

        If Ae_ON.Checked Then
            'On
            Return (180 * Math.Sin(_Kap_ * Math.PI / 180) * _Ae_ * _D_ * _ft_) / (Math.PI * _D_ * (180 / Math.PI * Math.Asin(_Ae_)))
        Else
            'Tangent
            Return (360 * Math.Sin(_Kap_ * Math.PI / 180) * _Ae_ * _D_ * _ft_) / (Math.PI * _D_ * (180 / Math.PI * Math.Acos(1 - 2 * _Ae_)))
        End If
    End Function

    'Werkelijke spaandikte  ### TURN & DILL ###
    Private Function Calc_T_Fnw(_fn_ As Double, _Kap_ As Double) As Double
        Return _fn_ * Math.Cos(_Kap_ * Math.PI / 180)
    End Function

    'Werkelijke senedediepte  ### TURN ###
    Private Function Calc_T_Apw(_ap_ As Double, _Kap_ As Double) As Double
        If _Kap_ = 90 Then
            Return 0
        Else
            Return _ap_ / Math.Cos(_Kap_ * Math.PI / 180)
        End If
    End Function

    'Verspaand volume ### MILL ###
    Private Function Calc_Q(_Ap_ As Double, _Ae_ As Double, _f_ As Double) As Double
        'MsgBox(_Ap_ & " " & _Ae_)
        Return _Ap_ * _Ae_ * _f_ / 1000
    End Function

    'Verspaand volume ### DRILL ###
    Private Function Calc_D_Q(_d_ As Double, _f_ As Double) As Double
        Return Math.PI * _d_ * _f_ / 1000
    End Function

    'Verspaand volume ### TURN ###
    Private Function Calc_T_Q(_Ap_ As Double, _Fn_ As Double, _V_ As Double) As Double
        Return _Ap_ * _Fn_ * _V_
    End Function

    'Spaandoorsnede  ### MILL ###
    Private Function Calc_A(_Ap_ As Double, _Hm_ As Double) As Double
        Return _Ap_ * _Hm_
    End Function

    'Spaandoorsnede  ### TURN ###
    Private Function Calc_T_A(_fn_ As Double, _d_ As Double) As Double
        Return _fn_ * _d_ / 2
    End Function

    'Spaandoorsnede  ### DRILL ###
    Private Function Calc_D_A(_fn_ As Double, _d_ As Double) As Double
        Return _fn_ * _d_ / 2
    End Function

    'Aantal effectieve tanden  ### MILL ###
    Private Function Calc_Ze(_Z_ As Double, _Ae_ As Double) As Double
        Return _Z_ / 2 * _Ae_
    End Function

    'Specifieke snijkracht bij spaanhoek Gamma  ### MILL & TURN ###
    Private Function Calc_Kcg(_Kc_ As Double, _gama_ As Double, _hm_ As Double, _mc_ As Double) As Double
        'Return _Kc_ * ((1 - _gama_ / 100) / Math.Pow(_hm_, -_mc_))
        Return _Kc_ * Math.Pow(_hm_, -_mc_) * (1 - (_gama_ / 100))
    End Function

    'Snijkracht totaal  ### MILL ###
    Private Function Calc_Ft(_Kcg_ As Double, _A_ As Double, _Ze_ As Double) As Double
        Return _Kcg_ * _A_ * _Ze_
    End Function

    'Snijkracht totaal  ### TURN ###
    Private Function Calc_T_Ft(_Kcg_ As Double, _A_ As Double) As Double
        Return _Kcg_ * _A_
    End Function

    'Snijkracht Radiaal  ### MILL ###
    Private Function Calc_Fr(_Ft_ As Double, _alfa_ As Double) As Double
        Return _Ft_ * Math.Cos(_alfa_ * Math.PI / 180)
    End Function

    'Snijkracht Radiaal  ### TURN ###
    Private Function Calc_T_Fr(_Fc_ As Double, _kap_ As Double) As Double
        Return _Fc_ * Math.Cos(_kap_ * Math.PI / 180)
    End Function

    'Snijkracht Axiaal  ### MILL ###
    Private Function Calc_Fa(_Ft_ As Double, _alfa_ As Double) As Double
        Return _Ft_ * Math.Sin(_alfa_ * Math.PI / 180)
    End Function

    'Snijkracht Radiaal  ### TURN ###
    Private Function Calc_T_Fa(_Fc_ As Double, _kap_ As Double) As Double
        Return _Fc_ * Math.Sin(_kap_ * Math.PI / 180)
    End Function

    'Vermogen Netto  ### TURN ###
    Private Function Calc_T_Pnet(_Kcg_ As Double, _V_ As Double, _A_ As Double) As Double
        Return _Kcg_ * (_A_ * _V_ / 60) / 1000
    End Function

    'Vermogen Netto  ### DRILL ###
    Private Function Calc_D_Pnet(_Fn_ As Double, _V_ As Double, _D_ As Double, _Kc_ As Double) As Double
        Return (_Fn_ * _V_ * _D_ * _Kc_) / (240 * 1000)
    End Function

    'Vermogen Netto  ### MILL ###
    Private Function Calc_Pnet(_Ap_ As Double, _Ae_ As Double, _F_ As Double, _Kc_ As Double) As Double
        Return (_Ap_ * _Ae_ * _F_ * _Kc_) / 60000000
    End Function

    'Koppel  ### MILL ###
    Private Function Calc_M(_Ft_ As Double, _D_ As Double) As Double
        Return _Ft_ * _D_ / 2000
    End Function

    'Koppel  ### TURN ###
    Private Function Calc_T_M(_Fc_ As Double, _D_ As Double) As Double
        Return _Fc_ * _D_ / 2000
    End Function

    '#############################################################################################################################
    Private Sub Z_TextChanged(sender As Object, e As EventArgs) Handles Z.TextChanged
        Calculate()
    End Sub

    Private Sub kappa_TextChanged(sender As Object, e As EventArgs) Handles kappa.TextChanged
        Calculate()
    End Sub

    Private Sub Aem_Enter(sender As Object, e As EventArgs) Handles Aem.Enter
        Ae.Text = ""
    End Sub

    Private Sub Ae_Enter(sender As Object, e As EventArgs) Handles Ae.Enter
        Aem.Text = ""
    End Sub

    Private Sub dril_kc_TextChanged(sender As Object, e As EventArgs) Handles dril_kc.TextChanged
        Calculate()
    End Sub

    Private Sub dril_d_TextChanged(sender As Object, e As EventArgs) Handles dril_d.TextChanged
        Calculate()
    End Sub

    Private Sub dril_z_TextChanged(sender As Object, e As EventArgs) Handles dril_z.TextChanged
        Calculate()
    End Sub

    Private Sub dril_kappa_TextChanged(sender As Object, e As EventArgs) Handles dril_kappa.TextChanged
        Calculate()
    End Sub

    Private Sub dril_gama_TextChanged(sender As Object, e As EventArgs) Handles dril_gama.TextChanged
        Calculate()
    End Sub

    Private Sub dril_v_TextChanged(sender As Object, e As EventArgs) Handles dril_v.TextChanged
        Calculate()
    End Sub

    Private Sub dril_fn_TextChanged(sender As Object, e As EventArgs) Handles dril_fn.TextChanged
        Calculate()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("K:\Systems\Visual Studio Applications\Visual Basic .NET\CuttingCalculator\_grob_g550t.png")
    End Sub

    Private Sub n_TextChanged(sender As Object, e As EventArgs) Handles n.TextChanged
        Calculate()
    End Sub

    Private Sub f_TextChanged(sender As Object, e As EventArgs) Handles f.TextChanged
        Calculate()
    End Sub

    Private Sub Toeren_label_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Toeren_label.LinkClicked
        Switch_N_V()
    End Sub

    Private Sub V_label_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles V_label.LinkClicked
        Switch_N_V()
    End Sub

    Private Sub alfa_TextChanged(sender As Object, e As EventArgs) Handles alfa.TextChanged
        Calculate()
    End Sub

    Private Sub gama_TextChanged(sender As Object, e As EventArgs) Handles gama.TextChanged
        Calculate()
    End Sub

    Private Sub PictureBox3_DoubleClick(sender As Object, e As EventArgs) Handles PictureBox3.DoubleClick
        System.Diagnostics.Process.Start("K:\Systems\Visual Studio Applications\Visual Basic .NET\CuttingCalculator\_kctable.png")
    End Sub

    Private Sub V_TextChanged(sender As Object, e As EventArgs) Handles V.TextChanged
        Calculate()
    End Sub

    Private Sub ft_TextChanged(sender As Object, e As EventArgs) Handles ft.TextChanged
        Calculate()
    End Sub

    Private Sub Ap_TextChanged(sender As Object, e As EventArgs) Handles Ap.TextChanged
        Calculate()
    End Sub

    Private Sub Aem_TextChanged(sender As Object, e As EventArgs) Handles Aem.TextChanged
        Calculate()
    End Sub

    Private Sub Ae_TextChanged(sender As Object, e As EventArgs) Handles Ae.TextChanged
        Calculate()
    End Sub

    Private Sub d_TextChanged(sender As Object, e As EventArgs) Handles d.TextChanged
        Calculate()
    End Sub

    Private Sub Kc0_TextChanged(sender As Object, e As EventArgs) Handles Kc0.TextChanged
        Calculate()
    End Sub

    Private Sub turn_kappa_TextChanged(sender As Object, e As EventArgs) Handles turn_kappa.TextChanged
        Calculate()
    End Sub

    Private Sub turn_Kc_TextChanged(sender As Object, e As EventArgs) Handles turn_Kc.TextChanged
        Calculate()
    End Sub

    Private Sub turn_d_TextChanged(sender As Object, e As EventArgs) Handles turn_d.TextChanged
        Calculate()
    End Sub

    Private Sub turn_mcfac_TextChanged(sender As Object, e As EventArgs) Handles turn_mcfac.TextChanged
        Calculate()
    End Sub

    Private Sub turn_gama_TextChanged(sender As Object, e As EventArgs) Handles turn_gama.TextChanged
        Calculate()
    End Sub

    Private Sub turn_V_TextChanged(sender As Object, e As EventArgs) Handles turn_V.TextChanged
        Calculate()
    End Sub

    Private Sub turn_Fn_TextChanged(sender As Object, e As EventArgs) Handles turn_Fn.TextChanged
        Calculate()
    End Sub

    Private Sub turn_Ap_TextChanged(sender As Object, e As EventArgs) Handles turn_Ap.TextChanged
        Calculate()
    End Sub

    Private Sub Ae_TAN_CheckedChanged(sender As Object, e As EventArgs) Handles Ae_TAN.CheckedChanged
        If Ae_TAN.Checked = True Then Ae_ON.Checked = False
        PictureBox1.Image = My.Resources.mill1
        PictureBox2.Image = My.Resources.mill20
        Calculate()
    End Sub

    Private Sub Ae_ON_CheckedChanged(sender As Object, e As EventArgs) Handles Ae_ON.CheckedChanged
        If Ae_ON.Checked = True Then Ae_TAN.Checked = False
        PictureBox1.Image = My.Resources.mill10
        PictureBox2.Image = My.Resources.mill2
        Calculate()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.Visible = False
        Panel3.Visible = False
    End Sub

    Private Sub kappa_Enter(sender As Object, e As EventArgs) Handles kappa.Enter
        Panel1.Visible = True
        Panel1.BringToFront()
        Tip.Image = CuttingCalculator.My.Resources.Resources.k
    End Sub

    Private Sub kappa_Leave(sender As Object, e As EventArgs) Handles kappa.Leave
        Panel1.Visible = False
        Calculate()
    End Sub
    Private Sub turn_kappa_Enter(sender As Object, e As EventArgs) Handles turn_kappa.Enter
        Panel3.Visible = True
        Panel3.BringToFront()
        Tip.Image = CuttingCalculator.My.Resources.Resources.turn_angle_k
    End Sub

    Private Sub turn_kappa_Leave(sender As Object, e As EventArgs) Handles turn_kappa.Leave
        Panel3.Visible = False
        Calculate()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CaptureScreen()
        printDocument1.Print()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CaptureScreen()
        Clipboard.SetDataObject(memoryImage)
    End Sub


    Private Sub CaptureScreen()
        Dim myGraphics As Graphics = Me.CreateGraphics()
        Dim s As Size = Me.Size
        memoryImage = New Bitmap(s.Width, s.Height, myGraphics)
        Dim memoryGraphics As Graphics = Graphics.FromImage(memoryImage)
        memoryGraphics.CopyFromScreen(Me.Location.X, Me.Location.Y, 0, 0, s)

    End Sub

    Private Sub printDocument1_PrintPage(ByVal sender As System.Object,
       ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles _
       printDocument1.PrintPage
        e.Graphics.DrawImage(memoryImage, 0, 0)
    End Sub


End Class
