import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  CommandBar,
  ICommandBarItemProps,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { IExternalUserManagerProps } from './IExternalUserManagerProps';
import { IExternalLibrary } from '../models/IExternalLibrary';
import { MockDataService } from '../services/MockDataService';
import styles from './ExternalUserManager.module.scss';

const ExternalUserManager: React.FC<IExternalUserManagerProps> = (props) => {
  const [libraries, setLibraries] = useState<IExternalLibrary[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedLibraries, setSelectedLibraries] = useState<IExternalLibrary[]>([]);
  const [selection] = useState(new Selection({
    onSelectionChanged: () => {
      setSelectedLibraries(selection.getSelection() as IExternalLibrary[]);
    }
  }));

  useEffect(() => {
    // Simulate loading data
    setTimeout(() => {
      const mockLibraries = MockDataService.getExternalLibraries();
      setLibraries(mockLibraries);
      setLoading(false);
    }, 1000);
  }, []);

  const columns: IColumn[] = [
    {
      key: 'name',
      name: 'Library Name',
      fieldName: 'name',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: IExternalLibrary) => (
        <Stack>
          <Text variant="medium" styles={{ root: { fontWeight: 'semibold' } }}>
            {item.name}
          </Text>
          <Text variant="small" styles={{ root: { color: '#666' } }}>
            {item.description}
          </Text>
        </Stack>
      )
    },
    {
      key: 'siteUrl',
      name: 'Site URL',
      fieldName: 'siteUrl',
      minWidth: 200,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IExternalLibrary) => (
        <Text variant="small" styles={{ root: { fontFamily: 'monospace' } }}>
          {item.siteUrl}
        </Text>
      )
    },
    {
      key: 'externalUsersCount',
      name: 'External Users',
      fieldName: 'externalUsersCount',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IExternalLibrary) => (
        <Text variant="medium" styles={{ root: { textAlign: 'center' } }}>
          {item.externalUsersCount}
        </Text>
      )
    },
    {
      key: 'owner',
      name: 'Owner',
      fieldName: 'owner',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'permissions',
      name: 'Permission Level',
      fieldName: 'permissions',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IExternalLibrary) => {
        const permissionColor = 
          item.permissions === 'Full Control' ? '#d73502' :
          item.permissions === 'Contribute' ? '#f7630c' : '#0078d4';
        
        return (
          <Text 
            variant="small" 
            styles={{ 
              root: { 
                color: permissionColor, 
                fontWeight: 'semibold',
                padding: '4px 8px',
                backgroundColor: `${permissionColor}15`,
                borderRadius: '4px'
              } 
            }}
          >
            {item.permissions}
          </Text>
        );
      }
    },
    {
      key: 'lastModified',
      name: 'Last Modified',
      fieldName: 'lastModified',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IExternalLibrary) => (
        <Text variant="small">
          {item.lastModified.toLocaleDateString()}
        </Text>
      )
    }
  ];

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'addLibrary',
      text: 'Add Library',
      iconProps: { iconName: 'Add' },
      onClick: () => {
        // Placeholder for add library functionality
        alert('Add Library functionality will be implemented');
      }
    },
    {
      key: 'removeLibrary',
      text: 'Remove',
      iconProps: { iconName: 'Delete' },
      disabled: selectedLibraries.length === 0,
      onClick: () => {
        // Placeholder for remove library functionality
        alert(`Remove functionality will be implemented for ${selectedLibraries.length} selected library(ies)`);
      }
    },
    {
      key: 'manageUsers',
      text: 'Manage Users',
      iconProps: { iconName: 'People' },
      disabled: selectedLibraries.length !== 1,
      onClick: () => {
        // Placeholder for manage users functionality
        if (selectedLibraries.length === 1) {
          alert(`Manage Users functionality will be implemented for: ${selectedLibraries[0].name}`);
        }
      }
    }
  ];

  const commandBarFarItems: ICommandBarItemProps[] = [
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        setLoading(true);
        setTimeout(() => {
          const mockLibraries = MockDataService.getExternalLibraries();
          setLibraries(mockLibraries);
          setLoading(false);
        }, 1000);
      }
    }
  ];

  return (
    <div className={styles.externalUserManager}>
      <Stack tokens={{ childrenGap: 20 }}>
        <Stack.Item>
          <Text variant="xxLarge" styles={{ root: { fontWeight: 'semibold' } }}>
            External User Manager
          </Text>
          <Text variant="medium" styles={{ root: { color: '#666', marginTop: '8px' } }}>
            Manage external users and shared libraries across SharePoint sites
          </Text>
        </Stack.Item>

        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.info}>
            This is a preview of the External User Manager. Use the action buttons to manage libraries and external user access.
          </MessageBar>
        </Stack.Item>

        <Stack.Item>
          <CommandBar
            items={commandBarItems}
            farItems={commandBarFarItems}
            ariaLabel="External Library Actions"
          />
        </Stack.Item>

        <Stack.Item>
          {loading ? (
            <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
              <Spinner size={SpinnerSize.large} label="Loading external libraries..." />
            </Stack>
          ) : (
            <DetailsList
              items={libraries}
              columns={columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selection={selection}
              selectionPreservedOnEmptyClick={true}
              selectionMode={SelectionMode.multiple}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
            />
          )}
        </Stack.Item>

        <Stack.Item>
          <Stack horizontal tokens={{ childrenGap: 15 }}>
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              Total Libraries: {libraries.length}
            </Text>
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              Selected: {selectedLibraries.length}
            </Text>
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              Total External Users: {libraries.reduce((sum, lib) => sum + lib.externalUsersCount, 0)}
            </Text>
          </Stack>
        </Stack.Item>
      </Stack>
    </div>
  );
};

export default ExternalUserManager;