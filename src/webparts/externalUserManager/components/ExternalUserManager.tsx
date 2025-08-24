import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
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
import { SharePointDataService } from '../services/SharePointDataService';
import { CreateLibraryModal } from './CreateLibraryModal';
import { DeleteLibraryModal } from './DeleteLibraryModal';
import { ManageUsersModal } from './ManageUsersModal';
import styles from './ExternalUserManager.module.scss';

const ExternalUserManager: React.FC<IExternalUserManagerProps> = (props) => {
  const [libraries, setLibraries] = useState<IExternalLibrary[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedLibraries, setSelectedLibraries] = useState<IExternalLibrary[]>([]);
  const [showCreateModal, setShowCreateModal] = useState<boolean>(false);
  const [showDeleteModal, setShowDeleteModal] = useState<boolean>(false);
  const [showManageUsersModal, setShowManageUsersModal] = useState<boolean>(false);
  const [operationMessage, setOperationMessage] = useState<{ message: string; type: MessageBarType } | null>(null);
  const [dataService] = useState(() => new SharePointDataService(props.context));
  
  const [selection] = useState(new Selection({
    onSelectionChanged: () => {
      setSelectedLibraries(selection.getSelection() as IExternalLibrary[]);
    }
  }));

  const loadLibraries = useCallback(async (useMockData: boolean = false) => {
    setLoading(true);
    setOperationMessage(null);
    
    try {
      let librariesData: IExternalLibrary[];
      
      if (useMockData) {
        // For development/fallback, use mock data
        librariesData = MockDataService.getExternalLibraries();
      } else {
        // Use real SharePoint API
        librariesData = await dataService.getExternalLibraries();
      }
      
      setLibraries(librariesData);
      
      if (!useMockData && librariesData.length === 0) {
        setOperationMessage({
          message: 'No external libraries found. Libraries with external users will appear here.',
          type: MessageBarType.info
        });
      }
    } catch (error) {
      console.error('Error loading libraries:', error);
      setOperationMessage({
        message: `Failed to load libraries: ${error.message}. Falling back to demo data.`,
        type: MessageBarType.warning
      });
      
      // Fallback to mock data on error
      setLibraries(MockDataService.getExternalLibraries());
    } finally {
      setLoading(false);
    }
  }, [dataService]);

  useEffect(() => {
    // Try to load real data first, fallback to mock data if needed
    loadLibraries(false);
  }, [loadLibraries]);

  const handleCreateLibrary = async (config: {
    title: string;
    description?: string;
    enableExternalSharing?: boolean;
  }): Promise<IExternalLibrary> => {
    try {
      const newLibrary = await dataService.createLibrary(config);
      return newLibrary;
    } catch (error) {
      console.error('Error creating library:', error);
      throw error;
    }
  };

  const handleLibraryCreated = (newLibrary: IExternalLibrary): void => {
    setLibraries(prev => [...prev, newLibrary]);
    setOperationMessage({
      message: `Library "${newLibrary.name}" created successfully.`,
      type: MessageBarType.success
    });
    
    // Clear message after 5 seconds
    setTimeout(() => setOperationMessage(null), 5000);
  };

  const handleDeleteLibrary = async (libraryId: string): Promise<void> => {
    try {
      await dataService.deleteLibrary(libraryId);
    } catch (error) {
      console.error('Error deleting library:', error);
      throw error;
    }
  };

  const handleLibrariesDeleted = (deletedIds: string[]): void => {
    setLibraries(prev => prev.filter(lib => deletedIds.indexOf(lib.id) === -1));
    selection.setAllSelected(false);
    
    const count = deletedIds.length;
    setOperationMessage({
      message: `${count} librar${count !== 1 ? 'ies' : 'y'} deleted successfully.`,
      type: MessageBarType.success
    });
    
    // Clear message after 5 seconds
    setTimeout(() => setOperationMessage(null), 5000);
  };

  const handleAddUser = async (libraryId: string, email: string, permission: 'Read' | 'Contribute' | 'Full Control'): Promise<void> => {
    try {
      await dataService.addExternalUserToLibrary(libraryId, email, permission);
      
      // Update the external users count for the library
      setLibraries(prev => prev.map(lib => 
        lib.id === libraryId 
          ? { ...lib, externalUsersCount: lib.externalUsersCount + 1 }
          : lib
      ));
      
    } catch (error) {
      console.error('Error adding user:', error);
      throw error;
    }
  };

  const handleBulkAddUsers = async (libraryId: string, emails: string[], permission: 'Read' | 'Contribute' | 'Full Control'): Promise<any> => {
    try {
      const results = await dataService.bulkAddExternalUsersToLibrary(libraryId, {
        emails,
        permission
      });
      
      // Count successful additions to update library count
      const successfulAdditions = results.filter(r => 
        r.status === 'success' || r.status === 'invitation_sent'
      ).length;
      
      // Update the external users count for the library
      if (successfulAdditions > 0) {
        setLibraries(prev => prev.map(lib => 
          lib.id === libraryId 
            ? { ...lib, externalUsersCount: lib.externalUsersCount + successfulAdditions }
            : lib
        ));
      }
      
      return results;
    } catch (error) {
      console.error('Error bulk adding users:', error);
      throw error;
    }
  };

  const handleRemoveUser = async (libraryId: string, userId: string): Promise<void> => {
    try {
      await dataService.removeExternalUserFromLibrary(libraryId, userId);
      
      // Update the external users count for the library
      setLibraries(prev => prev.map(lib => 
        lib.id === libraryId 
          ? { ...lib, externalUsersCount: Math.max(0, lib.externalUsersCount - 1) }
          : lib
      ));
      
    } catch (error) {
      console.error('Error removing user:', error);
      throw error;
    }
  };

  const handleGetUsers = async (libraryId: string) => {
    try {
      return await dataService.getExternalUsersForLibrary(libraryId);
    } catch (error) {
      console.error('Error getting users:', error);
      throw error;
    }
  };

  const handleSearchUsers = async (query: string) => {
    try {
      return await dataService.searchUsers(query);
    } catch (error) {
      console.error('Error searching users:', error);
      throw error;
    }
  };

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
      onClick: () => setShowCreateModal(true)
    },
    {
      key: 'removeLibrary',
      text: 'Remove',
      iconProps: { iconName: 'Delete' },
      disabled: selectedLibraries.length === 0,
      onClick: () => setShowDeleteModal(true)
    },
    {
      key: 'manageUsers',
      text: 'Manage Users',
      iconProps: { iconName: 'People' },
      disabled: selectedLibraries.length !== 1,
      onClick: () => {
        // Open manage users modal for the selected library
        if (selectedLibraries.length === 1) {
          setShowManageUsersModal(true);
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
        loadLibraries(false);
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
            Create and manage document libraries with external sharing capabilities. 
            Use PnPjs for SharePoint integration with Microsoft Graph API fallback.
          </MessageBar>
        </Stack.Item>

        {operationMessage && (
          <Stack.Item>
            <MessageBar messageBarType={operationMessage.type}>
              {operationMessage.message}
            </MessageBar>
          </Stack.Item>
        )}

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

      {/* Create Library Modal */}
      <CreateLibraryModal
        isOpen={showCreateModal}
        onClose={() => setShowCreateModal(false)}
        onLibraryCreated={handleLibraryCreated}
        onCreateLibrary={handleCreateLibrary}
      />

      {/* Delete Library Modal */}
      <DeleteLibraryModal
        isOpen={showDeleteModal}
        libraries={selectedLibraries}
        onClose={() => setShowDeleteModal(false)}
        onLibrariesDeleted={handleLibrariesDeleted}
        onDeleteLibrary={handleDeleteLibrary}
      />

      {/* Manage Users Modal */}
      <ManageUsersModal
        isOpen={showManageUsersModal}
        library={selectedLibraries.length === 1 ? selectedLibraries[0] : null}
        onClose={() => setShowManageUsersModal(false)}
        onAddUser={handleAddUser}
        onBulkAddUsers={handleBulkAddUsers}
        onRemoveUser={handleRemoveUser}
        onGetUsers={handleGetUsers}
        onSearchUsers={handleSearchUsers}
      />
    </div>
  );
};

export default ExternalUserManager;