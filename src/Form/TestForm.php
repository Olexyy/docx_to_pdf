<?php

namespace Drupal\docx_to_pdf\Form;

use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Drupal\docx_to_pdf\Factory\IOFactory;
use Drupal\file\Entity\File;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Settings;

/**
 * Class TestForm.
 *
 * @ingroup docx_to_pdf
 */
class TestForm extends FormBase {

  /**
   * Returns a unique string identifying the form.
   *
   * @return string
   *   The unique string identifying the form.
   */
  public function getFormId() {
    return 'docx_to_pdf_test';
  }

  /**
   * Form submission handler.
   *
   * @param array $form
   *   An associative array containing the structure of the form.
   * @param \Drupal\Core\Form\FormStateInterface $form_state
   *   The current state of the form.
   *
   * @throws \Exception
   */
  public function submitForm(array &$form, FormStateInterface $form_state) {

    $rendererName = $form_state->getValue('renderer');
    switch ($rendererName) {
      case Settings::PDF_RENDERER_DOMPDF:
        $rendererLibraryPath = DRUPAL_ROOT . '/../vendor/dompdf/dompdf';
        break;
      case Settings::PDF_RENDERER_TCPDF:
        $rendererLibraryPath = DRUPAL_ROOT . '/../vendor/tecnickcom/tcpdf';
        break;
      case Settings::PDF_RENDERER_MPDF:
        $rendererLibraryPath = DRUPAL_ROOT . '/../vendor/mpdf/mpdf';
        break;
      default:
        throw new \Exception('Unexpected value');
    }
    $form_file = $form_state->getValue('docx', 0);
    if (isset($form_file[0]) && !empty($form_file[0])) {
      $file = File::load($form_file[0]);
      $source = \Drupal::service('file_system')
        ->realpath($file->getFileUri());
      Settings::setPdfRendererPath($rendererLibraryPath);
      Settings::setPdfRendererName($rendererName);
      $phpWord = IOFactory::load($source);
      $this->export($phpWord, "sample_{$rendererName}.pdf", 'PDF', TRUE);
      // OR
      //$pdfWriter = IOFactory::createWriter($phpWord, 'PDF');
      //$pdfWriter->save('sample.pdf');
    }
  }

  /**
   * Save to file or download
   *
   * All exceptions should already been handled by the writers
   *
   * @param \PhpOffice\PhpWord\PhpWord $phpWord
   * @param string $filename
   * @param string $format
   * @param bool $download
   *
   * @return bool
   * @throws \PhpOffice\PhpWord\Exception\Exception
   */
  public function export(PhpWord $phpWord, $filename, $format = 'Word2007', $download = false)
  {
    $mime = array(
      'Word2007'  => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'ODText'    => 'application/vnd.oasis.opendocument.text',
      'RTF'       => 'application/rtf',
      'HTML'      => 'text/html',
      'PDF'       => 'application/pdf',
    );

    $writer = IOFactory::createWriter($phpWord, $format);

    if ($download === true) {
      header('Content-Description: File Transfer');
      header('Content-Disposition: attachment; filename="' . $filename . '"');
      header('Content-Type: ' . $mime[$format]);
      header('Content-Transfer-Encoding: binary');
      header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
      header('Expires: 0');
      $filename = 'php://output'; // Change filename to force download
    }

    $writer->save($filename);

    return true;
  }

  /**
   * Defines the settings form for Decision entities.
   *
   * @param array $form
   *   An associative array containing the structure of the form.
   * @param \Drupal\Core\Form\FormStateInterface $form_state
   *   The current state of the form.
   *
   * @return array
   *   Form definition array.
   */
  public function buildForm(array $form, FormStateInterface $form_state) {

    $form['form_wrapper']['#markup'] = $this->t('Test docx o pdf functionality.') . '</br>';
    $form['form_wrapper']['docx'] = [
      '#type' => 'managed_file',
      '#title' => $this->t('Docx file'),
      '#upload_location' => 'public://',
      '#upload_validators' => [
        'file_validate_extensions' => ['docx'],
      ],
      '#required' => TRUE,
    ];
    $form['form_wrapper']['renderer'] = [
      '#type' => 'radios',
      '#options' => [
        Settings::PDF_RENDERER_DOMPDF => Settings::PDF_RENDERER_DOMPDF,
        Settings::PDF_RENDERER_TCPDF =>  Settings::PDF_RENDERER_TCPDF,
        Settings::PDF_RENDERER_MPDF => Settings::PDF_RENDERER_MPDF,
      ],
      '#required' => TRUE,
      '#default_value' => Settings::PDF_RENDERER_TCPDF,
    ];
    $form['form_wrapper']['submit'] = [
      '#type' => 'submit',
      '#value' => $this->t('Submit'),
    ];

    return $form;
  }

}
