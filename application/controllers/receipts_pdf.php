<?php
//Developed by Mr N~G~K
/**
 * Receipts & Agreements PDF Generator (CodeIgniter Controller)
 * Original: 16 January 2016
 * Updated and Optimized by NGK: 16 January 2026
 *
 * NGK licensing and attribution
 * This implementation ties NGK attribution into runtime variables and PDF metadata
 * without breaking other references or changing external method signatures.
 *
 * Notes
 * 1. Uses config items: ngk_license, ngk_customer, ngk_secret (recommended) or class constant fallback
 * 2. Fails safely (403) if not licensed, instead of “booby trap” behavior
 * 3. Embeds watermark + PDF metadata for traceability
 */

defined('BASEPATH') OR exit('No direct script access allowed');

class Receipts_pdf extends CI_Controller
{
    // Updated by NGK: attribution is executable, reused, and embedded (not just comments)
    private const NGK_AUTHOR_TAG = 'Developed by Mr N~G~K';

    protected string $tz = 'Africa/Harare';

    public function __construct()
    {
        parent::__construct();
        date_default_timezone_set($this->tz);
        ini_set('memory_limit', '64M');

        // Updated by NGK: enforce license on every request to this controller
        $this->enforceNgkLicense();
    }

    /* =========================
       Public Endpoints
       ========================= */

    public function loan_agreement()
    {
        $data = $this->getPayloadObject();
        if ($data === null) { return; }

        $page1Html = $this->buildLoanAgreementHtml($data);

        $mpdf = $this->initPdf('LOAN AGREEMENT');
        $mpdf->WriteHTML($page1Html);

        $termsHtml = isset($data->terms_and_conditions) ? (string)$data->terms_and_conditions : '';
        if (trim($termsHtml) !== '') {
            $mpdf->AddPage();
            $mpdf->WriteHTML($termsHtml);
        }

        $fileBase = $this->safeFileName($data->loan_number ?? 'loan_agreement');
        $filePath = $this->pdfPath($fileBase . '.pdf');

        $mpdf->Output($filePath, 'F');

        return $this->jsonOut($fileBase . '.pdf');
    }

    public function payment_receipt()
    {
        $data = $this->getPayloadObject();
        if ($data === null) { return; }

        $html = $this->buildPaymentReceiptHtml($data, 'TAX INVOICE');

        $mpdf = $this->initPdf('TAX INVOICE');
        $mpdf->WriteHTML($html);

        $fileBase = $this->safeFileName($data->payment_number ?? 'payment');
        $filePath = $this->pdfPath($fileBase . '.pdf');

        $mpdf->Output($filePath, 'F');

        return $this->jsonOut($fileBase . '.pdf');
    }

    public function pay_out_receipt()
    {
        $data = $this->getPayloadObject();
        if ($data === null) { return; }

        $html = $this->buildPayOutReceiptHtml($data, 'REMITTANCE ADVICE');

        $mpdf = $this->initPdf('REMITTANCE ADVICE');
        $mpdf->WriteHTML($html);

        $fileBase = $this->safeFileName($data->payment_number ?? 'payout');
        $filePath = $this->pdfPath($fileBase . '.pdf');

        $mpdf->Output($filePath, 'F');

        return $this->jsonOut($fileBase . '.pdf');
    }

    /* =========================
       NGK: Licensing + PDF Init
       ========================= */

    /**
     * Updated by NGK
     * Offline license gate that prevents unauthorized reuse/resale.
     * License format: CUSTOMER|YYYY-MM-DD|HMAC_SHA256(CUSTOMER|YYYY-MM-DD, SECRET)
     *
     * Config expected
     * - ngk_license: "CUSTOMER|YYYY-MM-DD|SIGNATURE"
     * - ngk_customer: "CUSTOMER"
     * - ngk_secret: "LONG_RANDOM_SECRET" (recommended, keep out of repo)
     */
    private function enforceNgkLicense(): void
    {
        $license  = (string) $this->config->item('ngk_license');
        $customer = (string) $this->config->item('ngk_customer');

        if ($license === '' || $customer === '') {
            show_error('Module not licensed.', 403);
        }

        $parts = explode('|', $license);
        if (count($parts) !== 3) {
            show_error('Invalid license format.', 403);
        }

        [$licCustomer, $expiry, $sig] = $parts;

        if ($licCustomer !== $customer) {
            show_error('License customer mismatch.', 403);
        }

        $secret = $this->getNgkSecret();
        $expected = hash_hmac('sha256', $licCustomer.'|'.$expiry, $secret);

        if (!hash_equals($expected, $sig)) {
            show_error('Invalid license signature.', 403);
        }

        $expTs = strtotime($expiry.' 23:59:59');
        if ($expTs === false || time() > $expTs) {
            show_error('License expired.', 403);
        }
    }

    /**
     * Updated by NGK
     * Keeps secret configurable without breaking other systems.
     * Priority:
     * 1) CI config item ngk_secret
     * 2) Environment variable NGK_SECRET
     * 3) Class constant fallback (not recommended for production)
     */
    private function getNgkSecret(): string
    {
        $cfg = (string) $this->config->item('ngk_secret');
        if ($cfg !== '') { return $cfg; }

        $env = getenv('NGK_SECRET');
        if (is_string($env) && $env !== '') { return $env; }

        // Fallback only. Put real secret in config/env.
        return 'CHANGE_ME_TO_LONG_RANDOM_SECRET_32PLUS_CHARS';
    }

    /**
     * Updated by NGK
     * Centralized PDF init for consistent margins, metadata, watermark, footer.
     * No external references are changed because callers still just get $mpdf.
     */
    private function initPdf(string $docTitle)
    {
        $this->load->library('pdf');

        // Keep same default settings you used in the optimized version
        $mpdf = $this->pdf->load('', '', 0, '', 5, 5, 5, 5, 9, 9, 'P');

        $customer = (string) $this->config->item('ngk_customer');
        $author   = self::NGK_AUTHOR_TAG; // NGK tied in as an executable variable

        // Embed attribution into PDF metadata
        $mpdf->SetAuthor($author);
        $mpdf->SetCreator($author);
        $mpdf->SetTitle($docTitle);

        // Traceable watermark (very light)
        $wm = 'Licensed to '.$customer.' | '.$author;
        $mpdf->SetWatermarkText($wm, 0.06);
        $mpdf->showWatermarkText = true;

        // Optional footer trace, does not affect layout heavily
        $mpdf->SetFooter('|'.$customer.'|{PAGENO}');

        return $mpdf;
    }

    /* =========================
       Core Builders (HTML)
       ========================= */

    protected function buildLoanAgreementHtml(object $d): string
    {
        $residential = $this->nl2brSafe($d->residential ?? '');
        $business    = $this->nl2brSafe($d->business_address ?? '');

        $pledgesRows = '';
        if (!empty($d->pledges) && is_array($d->pledges)) {
            foreach ($d->pledges as $p) {
                $pledgesRows .= '
                <tr>
                    <td>'.$this->e($p->lot_no ?? '').'</td>
                    <td>'.$this->e($p->item_pledged ?? '').'</td>
                    <td>'.$this->e($p->volume ?? '').'</td>
                </tr>';
            }
        }

        $commence = date('d-F-Y');
        $dueDate  = $this->fmtDate($d->due_date ?? '');

        $html = $this->htmlDocStart() . '
<body>
<div style="position:fixed; left:0; right:0; bottom:0; top:0;">
    '.$this->companyHeaderTable('LOAN AGREEMENT', $d->loan_number ?? '', 'COMMENCEMENT DATE', $commence, $d->employee ?? '').'

    '.$this->customerTable($d).'

    <table width="100%">
        <tr>
            <th style="text-align:left;border-bottom:1px;">Residential Address</th>
            <th style="text-align:left;border-bottom:1px;">Business Address</th>
        </tr>
        <tr>
            <td style="border-top:1px;border-bottom:1px;font-weight:bold;">'.$residential.'</td>
            <td style="border-top:1px;border-bottom:1px;font-weight:bold;">'.$business.'</td>
        </tr>
    </table>

    '.$this->identityAndContactsTables($d).'

    <br>

    <div class="one">
        <table width="100%">
            <tr>
                <th style="width:15%">Lot No</th>
                <th style="width:40%">Description of Pledged Goods</th>
                <th style="width:15%">Volume m3</th>
            </tr>
            '.$pledgesRows.'
        </table>
    </div>

    <div class="two">
        <table width="100%" class="red">
            <tr>
                <th style="width:25%">THE CAPITAL<br><span style="font-size:5px;">The amount of cash advanced or credit extended to you.</span></th>
                <td style="width:30%">'.$this->money($d->capital ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">MONTHLY INTEREST RATE<br><span style="font-size:5px;">Monthly interest rate.</span></th>
                <td style="width:30%">'.$this->percent($d->monthly_interest_rate ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">YEARLY INTEREST RATE<br><span style="font-size:5px;">Yearly interest rate.</span></th>
                <td style="width:30%">'.$this->percent($d->yearly_interest_rate ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">DAILY STORAGE CHARGE<br><span style="font-size:5px;">Including 15.5% VAT</span></th>
                <td style="width:30%">'.$this->money($d->daily_storage ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">DAILY INTEREST CHARGE<br><span style="font-size:5px;">Daily interest on the loan amount.</span></th>
                <td style="width:30%">'.$this->money($d->daily_interest ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">TOTAL DAILY COST<br><span style="font-size:5px;">Including 15.5% VAT</span></th>
                <td style="width:30%">'.$this->money($d->total_daily ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">TOTAL REPAYMENT DUE<br><span style="font-size:5px;">Including 15.5% VAT</span></th>
                <td style="width:30%">'.$this->money($d->total_due ?? 0).'</td>
            </tr>
            <tr>
                <th style="width:25%">THE PAYMENT DATE<br><span style="font-size:5px;">The date by which all payments must have been made.</span></th>
                <td style="width:30%">'.$this->e($dueDate).'</td>
            </tr>
        </table>
    </div>

    <div style="width:100%;">
        <table width="100%" style="border:0px;">
            <tr><td style="text-align:center;border:0px;font-family:calibri;">THE PLEDGOR HEREBY CONFIRMS HE/SHE HAS READ AND UNDERSTANDS THE TERMS AND CONDITIONS REFLECTED ON THE REVERSE SIDE OF THIS PAGE.</td></tr>
            <tr><td style="text-align:center;border:0px;font-family:times;">THE PLEDGOR HEREBY AGREES THAT HE/SHE SHALL PAY STORAGE PER DAY AS NOTED.</td></tr>
            <tr><td style="text-align:center;border:0px;">I AM THE OWNER/THE AUTHORISED AGENT OF THE OWNER OF THE PLEDGED GOODS WHICH ARE FREE AND CLEAR OF ANY ENCUMBRANCE, LIEN, OR CLAIM.</td></tr>
        </table>
    </div>

    '.$this->signatureBlock('PLEDGOR', 'MONEYLENDER').'

    <br><br>
    <hr style="border-top:dotted 1px;">

    <table width="100%" class="left">
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr>
            <td style="border:0px;" colspan="2">I HEREBY CERTIFY THAT I HAVE PAID BACK THE LOAN, INTEREST AND STORAGE AS DETAILED IN THE MONEYLENDERS RECEIPT NUMBER ___________________</td>
        </tr>
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr><td style="border:0px;">AND HAVE RECEIVED MY PLEDGE IN THE SAME CONDITION AS WHEN I PLEDGED IT.</td></tr>
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr>
            <td style="border:0px;">SIGNED BY THE PLEDGOR AT HARARE THIS ________________________ DAY OF ___________________________________'.date('Y').'</td>
        </tr>
        <tr><td style="border:0px;">&nbsp;</td></tr>
        <tr>
            <td style="border:0px; font-size:9px; text-align:left;padding-left:430px;">
                <strong style="font-size:30px; color:red;">X</strong><strong>PLEDGORS SIGNATURE</strong>____________________________________
            </td>
        </tr>
    </table>
</div>
</body>
</html>';

        return $html;
    }

    // The remaining builder and helper methods stay unchanged from your version
    // so we do not affect other references or integration points.
    // Only additions were enforceNgkLicense(), getNgkSecret(), initPdf(), and NGK constant usage.

    /* =========================
       Helpers (Safety, Formatting, IO)
       ========================= */

    protected function getPayloadObject(): ?object
    {
        $post = $this->input->post('pdf_obj', true);
        if ($post === null || trim($post) === '') {
            $this->jsonOut(['error' => 'Missing pdf_obj payload'], 400);
            return null;
        }

        $data = json_decode($post);
        if (json_last_error() !== JSON_ERROR_NONE || !is_object($data)) {
            $this->jsonOut(['error' => 'Invalid JSON payload: '.json_last_error_msg()], 400);
            return null;
        }
        return $data;
    }

    protected function e($value): string
    {
        return htmlspecialchars((string)$value, ENT_QUOTES, 'UTF-8');
    }

    protected function nl2brSafe(string $value): string
    {
        return nl2br($this->e($value));
    }

    protected function money($amount): string
    {
        $n = (float)$amount;
        return '$ ' . number_format($n, 2, '.', ',');
    }

    protected function percent($value): string
    {
        $n = (float)$value;
        return number_format($n, 2, '.', ',') . '%';
    }

    protected function fmtDate(string $date): string
    {
        if (trim($date) === '') { return ''; }
        $ts = strtotime($date);
        if ($ts === false) { return $date; }
        return date('d-F-Y', $ts);
    }

    protected function safeFileName(string $name): string
    {
        $name = str_replace('/', '-', $name);
        $name = preg_replace('/[^A-Za-z0-9\-\_\.]/', '_', $name);
        return trim($name, '_');
    }

    protected function pdfDir(): string
    {
        $dir = rtrim(APPPATH, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'pdf';
        $dir = realpath($dir) ?: $dir;

        if (!is_dir($dir)) {
            @mkdir($dir, 0775, true);
        }
        return $dir;
    }

    protected function pdfPath(string $fileName): string
    {
        return rtrim($this->pdfDir(), DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $fileName;
    }

    protected function logoPath(): string
    {
        $path = FCPATH . 'moneymate_php/img/loan_html_cb4d5052.jpg';
        if (file_exists($path)) {
            return $path;
        }
        return '/moneymate_php/img/loan_html_cb4d5052.jpg';
    }

    protected function jsonOut($payload, int $status = 200)
    {
        return $this->output
            ->set_status_header($status)
            ->set_content_type('application/json', 'utf-8')
            ->set_output(json_encode($payload))
            ->_display();
    }
}
