# Developer Onboarding Guide

## Quick Start for New Developers

This guide will help you get started with the SharePoint External User Manager web part development.

## Prerequisites

### Required Software
- **Node.js**: Version 16.x or 18.x (Note: SPFx 1.18.2 doesn't support Node 20+)
- **npm**: Comes with Node.js
- **Git**: For version control
- **Visual Studio Code**: Recommended editor with SPFx extensions

### Recommended VS Code Extensions
- **SharePoint Framework Snippets**
- **TypeScript Importer**
- **ES7+ React/Redux/React-Native snippets**
- **Auto Rename Tag**
- **Bracket Pair Colorizer**

## Environment Setup

### 1. Node Version Management
If you have Node.js 20+ installed, use a version manager:

**Windows (nvm-windows):**
```bash
nvm install 18.17.1
nvm use 18.17.1
```

**macOS/Linux (nvm):**
```bash
nvm install 18.17.1
nvm use 18.17.1
```

### 2. SharePoint Framework CLI
```bash
npm install -g @microsoft/sharepoint-framework-yeoman-generator
npm install -g gulp-cli
```

### 3. Project Setup
```bash
# Clone the repository
git clone <repository-url>
cd sharepoint-external-user-manager

# Install dependencies
npm install

# Start development server
npm run serve
```

## Project Architecture

### Folder Structure
```
src/webparts/externalUserManager/
├── components/
│   ├── ExternalUserManager.tsx          # Main React component
│   ├── ExternalUserManager.module.scss  # Component styles
│   └── IExternalUserManagerProps.ts     # Component props interface
├── models/
│   └── IExternalLibrary.ts             # TypeScript interfaces
├── services/
│   └── MockDataService.ts              # Mock data provider
├── loc/                                 # Localization files
└── ExternalUserManagerWebPart.ts       # SPFx web part entry point
```

### Key Technologies
- **SharePoint Framework (SPFx)**: 1.18.2
- **React**: 17.0.1 with TypeScript
- **Fluent UI**: 8.x (Microsoft's design system)
- **SCSS Modules**: For component styling
- **Gulp**: Build system

## Development Workflow

### 1. Local Development
```bash
# Start the dev server (auto-reload enabled)
npm run serve

# The web part will be available at:
# https://localhost:4321/temp/manifests.js
```

### 2. Testing in SharePoint
1. Open SharePoint Online
2. Add `?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js` to any SharePoint page URL
3. Add the web part to the page

### 3. Building for Production
```bash
# Build the solution
npm run build

# Package for deployment
npm run package-solution

# The .sppkg file will be in sharepoint/solution/
```

## Code Structure Guide

### Component Development
The main component (`ExternalUserManager.tsx`) follows these patterns:

```typescript
// State management with React hooks
const [libraries, setLibraries] = useState<IExternalLibrary[]>([]);
const [loading, setLoading] = useState<boolean>(true);

// Fluent UI components
import { DetailsList, CommandBar, Stack } from '@fluentui/react';

// SCSS modules for styling
import styles from './ExternalUserManager.module.scss';
```

### Adding New Features

#### 1. Add New Mock Data
Edit `src/webparts/externalUserManager/services/MockDataService.ts`:

```typescript
export class MockDataService {
  public static getNewFeatureData(): YourDataType[] {
    return [
      // Your mock data here
    ];
  }
}
```

#### 2. Create New Models
Add interfaces in `src/webparts/externalUserManager/models/`:

```typescript
export interface IYourNewModel {
  id: string;
  name: string;
  // Other properties
}
```

#### 3. Add New Components
Create new React components in the `components/` folder following the existing pattern.

### Styling Guidelines
- Use SCSS modules for component-specific styles
- Follow Fluent UI design tokens when possible
- Use responsive design patterns (`@media` queries)
- Maintain consistent spacing using Fluent UI `Stack` components

```scss
.yourComponent {
  padding: 20px;
  background-color: #ffffff;
  
  @media (max-width: 768px) {
    padding: 10px;
  }
}
```

## Debugging Tips

### Common Issues
1. **Node Version**: Ensure you're using Node 16.x or 18.x
2. **Certificate Issues**: Trust the local dev certificate
3. **CORS Errors**: Use SharePoint's debug mode

### Debugging in Browser
1. Open Developer Tools (F12)
2. Use React Developer Tools extension
3. Console logs are available in browser console
4. Source maps are enabled for TypeScript debugging

## Backend Integration Plan

The current implementation uses mock data. To integrate with real SharePoint APIs:

### 1. Replace MockDataService
```typescript
import { SPHttpClient } from '@microsoft/sp-http';

export class SharePointDataService {
  constructor(private context: WebPartContext) {}
  
  public async getExternalLibraries(): Promise<IExternalLibrary[]> {
    // Use this.context.spHttpClient for API calls
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1
    );
    return response.json();
  }
}
```

### 2. Update Component Usage
```typescript
// Replace MockDataService with SharePointDataService
const dataService = new SharePointDataService(props.context);
const libraries = await dataService.getExternalLibraries();
```

## Testing Strategy

### Manual Testing Checklist
- [ ] Component loads without errors
- [ ] Mock data displays correctly
- [ ] Action buttons show appropriate alerts
- [ ] Selection functionality works
- [ ] Responsive design works on mobile
- [ ] Loading states display properly

### Future Automated Testing
Consider adding:
- Unit tests with Jest
- React component testing with React Testing Library
- E2E tests with Playwright

## Deployment Process

### 1. Development to Staging
```bash
npm run build
npm run package-solution
```

### 2. Upload to SharePoint App Catalog
1. Go to SharePoint Admin Center
2. Navigate to More Features > Apps > App Catalog
3. Upload the `.sppkg` file from `sharepoint/solution/`

### 3. Deploy to Sites
1. Go to Site Contents > Add an App
2. Find your web part and add it
3. Add to pages using the web part picker

## Next Development Priorities

1. **Backend Integration** (High Priority)
   - Replace MockDataService with SharePoint REST API calls
   - Implement proper error handling

2. **User Management Features** (Medium Priority)
   - Add user details modal/panel
   - Implement invite/remove user functionality

3. **Library Management** (Medium Priority)
   - Add library creation modal
   - Implement library deletion with confirmation

4. **Enhanced UI** (Low Priority)
   - Add filtering and search
   - Implement pagination for large datasets
   - Add bulk operations

## Support and Resources

### Official Documentation
- [SharePoint Framework Overview](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Fluent UI React Components](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- [SPFx API Reference](https://docs.microsoft.com/en-us/javascript/api/sp-webpart-base/)

### Community Resources
- [SharePoint Stack Exchange](https://sharepoint.stackexchange.com/)
- [PnP Community](https://pnp.github.io/)
- [SPFx Samples](https://github.com/pnp/sp-dev-fx-webparts)

## Getting Help

1. Check the existing documentation and code comments
2. Review the GitHub issues for similar problems
3. Ask questions in the team chat or create a GitHub issue
4. Consult the SharePoint community resources listed above

Happy coding! 🚀