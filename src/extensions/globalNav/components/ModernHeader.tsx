import * as React from 'react';
import { FC, useState, useEffect } from 'react';
import { SPFI } from '@pnp/sp';
import { INavLink, ICommandBarItemProps, IContextualMenuItem, CommandBar, Stack } from 'office-ui-fabric-react';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";

export interface IModernHeaderProps {
	sp: SPFI;
}

const ModernHeader: FC<IModernHeaderProps> = (props) => {
	const { sp } = props;
	const [loading, setLoading] = useState<boolean>(true);
	const [menuItems, setMenuItems] = useState<ICommandBarItemProps[]>(undefined);
	const [selMenu, setSelMenu] = useState<string>(undefined);

	const getActiveMenuItems = async (): Promise<any[]> => {
		const filQuery = `IsActive eq 1 and Position eq 'Top'`;
		return await sp.web.lists.getByTitle('Menus').items
			.select('ID', 'Title', 'PageUrl', 'IconName', 'Sequence', 'IsParent', 'ParentMenu/Id', 'ParentMenu/Title', 'IsActive')
			.expand('ParentMenu')
			.filter(filQuery)();
	};

	const _loadTopNavigation = async (): Promise<void> => {
		const menuItems: any[] = await getActiveMenuItems();
		console.log("Top Menu items: ", menuItems);
		if (menuItems.length > 0) {
			const navLinks: IContextualMenuItem[] = [];
			const navLink: ICommandBarItemProps[] = [];
			if (menuItems.length > 0) {
				const fil = menuItems.filter((mi: any) => mi.IsParent);
				if (fil && fil.length > 0) {
					fil.map((item: any) => {
						const subMenus: any[] = menuItems.filter((smi: any) => !smi.IsParent && smi.ParentMenu?.Id === item.ID);
						let navsubLink: IContextualMenuItem[] = [];
						if (subMenus && subMenus.length > 0) {
							subMenus.map((item: any) => {
								if (item.PageUrl?.Url.toLowerCase() === (window.location.origin + window.location.pathname).toLowerCase())
									setSelMenu(item.ID.toString());
								navsubLink.push({
									key: item.ID.toString(),
									text: item.Title,
									url: item.PageUrl?.Url,
									expandAriaLabel: item.Title,
									iconProps: { iconName: item.IconName },
									onClick: () => { window.location.href = item.PageUrl?.Url }
								});
							});
							if (item.PageUrl?.Url.toLowerCase() === (window.location.origin + window.location.pathname).toLowerCase())
								setSelMenu(item.ID.toString());
							navLinks.push({
								key: item.ID.toString(),
								text: item.Title,
								url: item.PageUrl?.Url,
								iconProps: { iconName: item.IconName },
								subMenuProps: { items: navsubLink },
								onClick: () => { window.location.href = item.PageUrl?.Url }
							});
						} else {
							if (item.PageUrl?.Url.toLowerCase() === (window.location.origin + window.location.pathname).toLowerCase())
								setSelMenu(item.ID.toString());
							navLinks.push({
								key: item.ID.toString(),
								text: item.Title,
								url: item.PageUrl?.Url,
								iconProps: { iconName: item.IconName },
								onClick: () => { window.location.href = item.PageUrl?.Url }
							});
						}
						navsubLink = [];
					});
				}
				setMenuItems(navLinks);
			}
		}
	};

	useEffect(() => {
		(async () => {
			await _loadTopNavigation();
		})();
	}, []);

	return (
		<div>
			<CommandBar
				items={menuItems}
			/>
		</div>
	)
};

export default ModernHeader;
