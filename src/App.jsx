import React from 'react';
import {
  makeStyles,
  shorthands,
  tokens,
  Body1,
  Body1Strong,
  Link,
  Switch,
  Button,
  Text,
  InfoButton,
  OverlayDrawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  DrawerFooter,
  useRestoreFocusSource,
  useRestoreFocusTarget,
  SpinButton,
  Label,
} from '@fluentui/react-components';
import { Settings24Regular, Dismiss24Regular } from '@fluentui/react-icons';
import { InfoLabel } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  main: {
    display: 'grid',
    justifyContent: 'flex-start',
    gridRowGap: tokens.spacingVerticalXXL,
  },
  header: {
    ...shorthands.padding('16px'),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  content: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    overflowY: 'auto',
    flex: 1,
    padding: '20px',
  },
  description: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground3,
    marginTop: tokens.spacingVerticalXS,
  },
  settingsContainer: {
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius('8px'),
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  settingRow: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '12px 16px',
  },
  settingHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: 0,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase700,
    fontWeight: tokens.fontWeightSemibold,
  },
  settingTitle: {
    flex: 1,
  },
  description: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    marginTop: 0,
    lineHeight: tokens.lineHeightBase200,
  },
  tooltip: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    maxWidth: '300px',
    ...shorthands.padding(tokens.spacingVerticalM),
    '& a': {
      marginTop: tokens.spacingVerticalXS,
      color: tokens.colorBrandForeground1,
      textDecorationLine: 'none',
      ':hover': {
        textDecorationLine: 'underline',
      },
    },
  },

  footer: {
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

const App = () => {
  const styles = useStyles();
  const [isOpen, setIsOpen] = React.useState(false);
  const [computerUseEnabled, setComputerUseEnabled] = React.useState(false);
  const [windowsSessionEnabled, setWindowsSessionEnabled] = React.useState(true);
  const [maintenanceModeEnabled, setMaintenanceModeEnabled] = React.useState(false);
  const [queuePrioritizationEnabled, setQueuePrioritizationEnabled] = React.useState(true);
  const [autoAllocationEnabled, setAutoAllocationEnabled] = React.useState(true);
  const [manualAllocationLimit, setManualAllocationLimit] = React.useState(0);
  
  const machineLimit = 5;
  const totalAvailableBots = 20;
  const availableBots = totalAvailableBots - manualAllocationLimit;
  
  // Setup focus management
  const restoreFocusTargetAttributes = useRestoreFocusTarget();
  const restoreFocusSourceAttributes = useRestoreFocusSource();

  return (
    <div className={styles.root}>
      <Button
        {...restoreFocusTargetAttributes}
        appearance="primary"
        icon={<Settings24Regular />}
        onClick={() => setIsOpen(true)}
      >
        Open Settings
      </Button>

      <OverlayDrawer
        {...restoreFocusSourceAttributes}
        position="end"
        open={isOpen}
        onOpenChange={(_, { open }) => setIsOpen(open)}
        size="medium"
        style={{ overflow: 'visible' }}
      >
        <DrawerHeader className={styles.header}>
          <DrawerHeaderTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={() => setIsOpen(false)}
              />
            }
          >
            Settings
          </DrawerHeaderTitle>
        </DrawerHeader>
        <DrawerBody style={{ display: 'flex', flexDirection: 'column', height: '100%', padding: 0 }}>
          <div className={styles.content}>
            {/* General Section */}
            <div className={styles.section}>
              <Body1Strong>General</Body1Strong>
              <div className={styles.settingsContainer}>
                {/* Enable Computer Use Setting */}
                <div 
                  className={styles.settingRow} 
                  style={{
                    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
                    paddingBottom: tokens.spacingVerticalM,
                  }}
                >
                  <div className={styles.settingHeader}>
                    <InfoLabel
                      className={styles.settingTitle}
                      style={{ fontWeight: 600 }}
                      info={
                        <>
                          Enable this setting to allow Microsoft Copilot Studio to interact with websites and desktop applications through automated clicking and typing. This machine will be dedicated to computer interactions, and other automation settings will be disabled. Disable this setting if you want to use this machine for standard desktop flows without computer interactions.
                          <br /><br />
                          Default: Off
                        </>
                      }
                      label="Enable computer use"
                    />
                    <Switch
                      checked={computerUseEnabled}
                      onChange={(e) => setComputerUseEnabled(e.target.checked)}
                    />
                  </div>
                  <Text className={styles.description}>
                    Allows Microsoft Copilot agents to perform tasks in websites and desktop applications. Turn this setting off if you want to use this machine for desktop flows.
                  </Text>
                </div>

                {/* Reuse Windows Session Setting */}
                <div className={styles.settingRow}>
                  <div className={styles.settingHeader}>
                    <InfoLabel
                      className={styles.settingTitle}
                      style={{ fontWeight: 600 }}
                      info={
                        <>
                          When enabled, the machine will reuse existing Windows sessions for attended desktop flow runs.
                          <br /><br />
                          Default: On
                          <br /><br />
                          <Link
                            href="https://learn.microsoft.com/en-us/power-automate/desktop-flows/run-desktop-flow#reuse-windows-session"
                            target="_blank"
                          >
                            Learn more about Windows session reuse
                          </Link>
                        </>
                      }
                      label="Reuse Windows session"
                    />
                    <Switch
                      checked={windowsSessionEnabled}
                      disabled={computerUseEnabled}
                      onChange={(e) => setWindowsSessionEnabled(e.target.checked)}
                    />
                  </div>
                  <Body1 className={styles.description}>
                    Reuse existing Windows sessions for attended desktop flow runs
                  </Body1>
                </div>

                {/* Maintenance Mode Setting */}
                <div className={styles.settingRow}>
                  <div className={styles.settingHeader}>
                    <InfoLabel
                      className={styles.settingTitle}
                      style={{ fontWeight: 600 }}
                      info={
                        <>
                          When enabled, the machine will be in maintenance mode and will not accept new desktop flow runs.
                          <br /><br />
                          Default: Off
                          <br /><br />
                          <Link
                            href="https://learn.microsoft.com/en-us/power-automate/desktop-flows/manage-machines#maintenance-mode"
                            target="_blank"
                          >
                            Learn more about maintenance mode
                          </Link>
                        </>
                      }
                      label="Enable maintenance mode"
                    />
                    <Switch
                      checked={maintenanceModeEnabled}
                      disabled={computerUseEnabled}
                      onChange={(e) => setMaintenanceModeEnabled(e.target.checked)}
                    />
                  </div>
                </div>

                {/* Extended Queue Prioritization Setting */}
                <div className={styles.settingRow}>
                  <div className={styles.settingHeader}>
                    <InfoLabel
                      className={styles.settingTitle}
                      style={{ fontWeight: 600 }}
                      info={
                        <>
                          When enabled, lower priority flows may be delayed to ensure higher priority flows are not blocked.
                          <br /><br />
                          Default: On
                          <br /><br />
                          <Link
                            href="https://learn.microsoft.com/en-us/power-automate/desktop-flows/monitor-desktop-flow-queues#extended-queue-prioritization"
                            target="_blank"
                          >
                            Learn more about queue prioritization
                          </Link>
                        </>
                      }
                      label="Extended queue prioritization"
                    />
                    <Switch
                      checked={queuePrioritizationEnabled}
                      disabled={computerUseEnabled}
                      onChange={(e) => setQueuePrioritizationEnabled(e.target.checked)}
                    />
                  </div>
                </div>
              </div> { /* End General settingsContainer */}
            </div> { /* End General section */}

            {/* Unattended Bots Section */}
            <div className={styles.section} style={{ marginTop: '16px' }}>
              <Body1Strong>Unattended bots</Body1Strong>
              <div className={styles.settingsContainer}>
                <div className={styles.settingRow}>
                  <div className={styles.settingHeader}>
                    <InfoLabel
                      className={styles.settingTitle}
                      info={
                        <>
                          When enabled, bots will be automatically allocated based on workload demands, ignoring manual allocation settings.
                          <br /><br />
                          Default: On
                        </>
                      }
                      label="Auto-allocation"
                    />
                    <Switch
                      checked={autoAllocationEnabled}
                      onChange={(e) => setAutoAllocationEnabled(e.target.checked)}
                    />
                  </div>
                  <Body1 className={styles.description}>
                    Dynamically allocate bots based on workload
                  </Body1>

                  {/* Manual Allocation Controls (Conditional) */}
                  {!autoAllocationEnabled && (
                    <div style={{ 
                      marginTop: '16px', 
                      marginLeft: '-16px',
                      marginRight: '-16px',
                      borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
                      padding: '16px 16px 0'
                    }}>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                          <InfoLabel
                            className={styles.settingTitle}
                            label={<Body1Strong>Manual allocation</Body1Strong>}
                            info={
                              <>
                                Set the maximum number of bots that can run on this machine.
                                <br /><br />
                                Default: 0
                              </>
                            }
                          />
                          <SpinButton
                            value={manualAllocationLimit}
                            min={0}
                            max={machineLimit}
                            onChange={(e, data) => setManualAllocationLimit(data.value)}
                            style={{ width: '80px' }}
                          />
                        </div>
                        <div
                          style={{
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '8px',
                            padding: '12px',
                            backgroundColor: tokens.colorNeutralBackground2,
                            borderRadius: '4px',
                          }}
                        >
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                            <Body1>Machine limit</Body1>
                            <Body1Strong style={{ color: tokens.colorNeutralForeground3 }}>{machineLimit}</Body1Strong>
                          </div>
                          <div
                            style={{
                              height: '1px',
                              backgroundColor: tokens.colorNeutralStroke2,
                              margin: '4px 0',
                            }}
                          />
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                            <Body1>Available bots in environment</Body1>
                            <Body1Strong style={{ color: tokens.colorBrandForeground1 }}>{availableBots}</Body1Strong>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div> { /* End Auto-allocation settingRow */}
              </div> { /* End Unattended bots settingsContainer */}
            </div> { /* End Unattended bots section */}
          </div> { /* End content */}
        </DrawerBody>
        <DrawerFooter className={styles.footer}>
          <Button appearance="primary">Save</Button>
          <Button appearance="secondary">Cancel</Button>
        </DrawerFooter>
      </OverlayDrawer>
    </div>
  );
};

export default App;
