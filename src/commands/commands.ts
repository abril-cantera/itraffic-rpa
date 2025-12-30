/**
 * Ribbon button command handlers
 */

Office.onReady(() => {
  console.log('📌 Commands loaded');
});

/**
 * Show taskpane when ribbon button is clicked
 */
function showTaskpane(event: Office.AddinCommands.Event): void {
  Office.addin.showAsTaskpane();
  event.completed();
}

/**
 * Quick action: Check classification status
 */
async function checkStatus(event: Office.AddinCommands.Event): Promise<void> {
  try {
    // This is a quick action that doesn't open the taskpane
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (item) {
      // Show notification
      mailbox.item?.notificationMessages.addAsync('status-check', {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: '✅ Email Classifier is active',
        icon: 'icon-16',
        persistent: false,
      });
    }
  } catch (error) {
    console.error('Error checking status:', error);
  } finally {
    event.completed();
  }
}

/**
 * Classify current email manually
 */
async function classifyEmail(event: Office.AddinCommands.Event): Promise<void> {
  console.log('🎯 classifyEmail called');
  
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (!item) {
      console.error('❌ No email item selected');
      event.completed();
      return;
    }

    // Show processing notification
    mailbox.item?.notificationMessages.addAsync('classify-processing', {
      type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
      message: '🤖 Clasificando email...',
    });

    // Get userId from Office Roaming Settings (shared across contexts)
    const settings = Office.context.roamingSettings;
    const userId = settings.get('email-classifier-userId');
    
    if (!userId) {
      console.error('❌ User not registered');
      mailbox.item?.notificationMessages.replaceAsync('classify-processing', {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: '⚠️ Debes activar el clasificador primero. Abre el panel lateral.',
      });
      event.completed();
      return;
    }
    
    console.log('👤 Using userId:', userId);

    // Get email ID
    const emailId = item.itemId;
    if (!emailId) {
      console.error('❌ No email ID available');
      mailbox.item?.notificationMessages.replaceAsync('classify-processing', {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: '⚠️ Email no disponible para clasificación',
      });
      event.completed();
      return;
    }

    console.log('📧 Classifying email:', emailId);

    // Call backend API
    const apiBaseUrl = 'https://app-itraffic-rpa.whiteflower-4df565a8.eastus2.azurecontainerapps.io';
    const response = await fetch(`${apiBaseUrl}/api/classify`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        userId: userId,
        messageId: emailId,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      // Backend returns { error: '...', details: '...' }
      throw new Error(error.error || error.message || 'Error al clasificar');
    }

    const result = await response.json();
    console.log('✅ Classification result:', result);

    // Show success notification
    mailbox.item?.notificationMessages.replaceAsync('classify-processing', {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `✅ Clasificado como: ${result.categories.join(', ')}`,
      icon: 'icon-16',
      persistent: false,
    });

  } catch (error: any) {
    console.error('❌ Error classifying email:', error);
    Office.context.mailbox.item?.notificationMessages.replaceAsync('classify-processing', {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `⚠️ Error: ${error?.message || 'Error desconocido'}`,
      // persistent: false is not supported for ErrorMessage
    });
  } finally {
    event.completed();
  }
}

// Register functions for ribbon buttons
(Office.actions as any).associate('showTaskpane', showTaskpane);
(Office.actions as any).associate('checkStatus', checkStatus);
(Office.actions as any).associate('classifyEmail', classifyEmail);
